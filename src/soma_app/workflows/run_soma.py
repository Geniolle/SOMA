from __future__ import annotations

import logging
import os
import re
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional

from soma_app.automation.adapters import EntradasSaidasAdapter, TransferenciasAdapter
from soma_app.automation.api.iv_api import EntradasSaidasApi, TransferenciasApi
from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.automation.pages.login_page import LoginPage
from soma_app.automation.pages.transferencias_page import TransferenciasPage
from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra import report as legacy_report
from soma_app.infra.env import env_bool, env_str
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.sheets_client import SheetsClient
from soma_app.infra.soma_api_client import SomaApiClient
from soma_app.infra.trace import log_kv, new_run_id, step
from soma_app.infra.webdriver_factory import (
    WebDriverFactory,
    get_chrome_version,
    get_chromedriver_info,
    unwrap_webdriver,
)
from soma_app.workflows.contaordem_writer import (
    mark_row_doc_soma,
    mark_row_duplicate,
    mark_row_error,
    mark_row_ok,
    unlock_still_processing,
)
from soma_app.workflows.process_caixas_bancos import atualizar_caixas_bancos
from soma_app.workflows.process_contaordem import (
    STATUS_COL_DEFAULT,
    SheetsTable,
    preprocess_contaordem,
)
from soma_app.workflows.process_soma import atualizar_sheet_soma

logger = logging.getLogger(__name__)

try:
    from dotenv import load_dotenv  # type: ignore

    load_dotenv()
except Exception:
    pass


def _is_pending_doc_exception(e: Exception) -> bool:
    try:
        tb = traceback.extract_tb(e.__traceback__)
        fnames = {f.name for f in tb}
        return "_go_back_to_list_best_effort" in fnames or "_search_doc_id" in fnames
    except Exception:
        return False


def _safe_err(e: Exception) -> str:
    s = str(e).strip().replace("\n", " ")
    return s[:180] if s else type(e).__name__


def _norm_status(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").strip().upper())


def _should_recover_doc(raw_row: Dict[str, Any]) -> bool:
    # STATUS=PENDENTE_DOC é o único sinal confiável: DOC. SOMA pode ter sido
    # sobrescrito para "Em processamento" (lock) ou "" (unlock_still_processing)
    # entre o run que criou o documento e este reprocessamento.
    status_value = raw_row.get(STATUS_COL_DEFAULT, "")
    return _norm_status(status_value) == "PENDENTEDOC"


def _load_settings() -> Any:
    from soma_app.config.settings import Settings

    # run_soma.py -> workflows -> soma_app -> src -> <project_root>
    project_root = Path(__file__).resolve().parents[3]

    env_file = (os.getenv("ENV_FILE") or "").strip()
    if env_file:
        env_path = Path(env_file)
        if not env_path.is_absolute():
            # 1) tenta relativo ao cwd (comportamento padrão)
            cwd_candidate = Path.cwd() / env_path
            if cwd_candidate.exists():
                return Settings.from_env(env_file=str(cwd_candidate))

            # 2) tenta relativo à raiz do projeto
            project_candidate = project_root / env_path
            if project_candidate.exists():
                return Settings.from_env(env_file=str(project_candidate))
        elif env_path.exists():
            return Settings.from_env(env_file=str(env_path))

    # fallback robusto: procura deploy/.env na raiz do projeto
    default_env = project_root / "deploy" / ".env"
    if default_env.exists():
        return Settings.from_env(env_file=str(default_env))

    return Settings.from_env(env_file=None)


def _sheet_name(settings: Any) -> str:
    env_ws = (os.getenv("SHEET_CONTAORDEM") or "").strip()
    if env_ws:
        return env_ws
    return (os.getenv("SHEET_NAME") or os.getenv("SHEET") or "TESTE_CONTAORDEM").strip()


def _run_post_processes(
    *,
    settings: Any,
    bundle: Any | None,
    run_caixas_bancos: bool,
) -> None:
    if bundle is None:
        logger.warning("Pós-processos ignorados: browser indisponível.")
        return

    caixas_ok = True

    if run_caixas_bancos:
        with step(logger, "run.post.caixas_bancos"):
            try:
                sheets_caixas = SheetsClient(settings)
                atualizar_caixas_bancos(sheets_caixas, bundle.a, settings)
            except Exception:
                caixas_ok = False
                logger.exception("Falha no processo Caixas/Bancos.")

    if not caixas_ok:
        logger.warning("Processo de atualização da sheet SOMA não será executado porque Caixas/Bancos falhou.")
        return

    with step(logger, "run.post.atualizar_sheet_soma"):
        try:
            sheets_soma = SheetsClient(settings)
            atualizar_sheet_soma(sheets_soma, bundle.a, settings)
        except Exception:
            logger.exception("Falha no processo de atualização da sheet SOMA.")


def _bootstrap_api_session(
    api_client: SomaApiClient,
    *,
    session_token: str,
    has_api_credentials: bool,
    auth_preference: str,
    close_on_exit: bool,
) -> None:
    preference = (auth_preference or "").strip().lower()
    if preference not in {"login_first", "token_first"}:
        preference = "token_first"

    token = (session_token or "").strip()
    prefer_login = preference == "login_first"

    if token and not prefer_login:
        api_client.use_session_token(token, close_on_exit=close_on_exit)
        return

    if has_api_credentials:
        try:
            api_client.open_session()
            return
        except Exception:
            if token:
                logger.warning(
                    "Falha ao abrir nova sessao na API com SOMA_API_LOGIN. "
                    "Usando SOMA_SESSION_TOKEN como fallback."
                )
                api_client.use_session_token(token, close_on_exit=close_on_exit)
                return
            raise

    if token:
        api_client.use_session_token(token, close_on_exit=close_on_exit)
        return

    api_client.open_session()


@dataclass
class RunState:
    """
    Acumula bundle/api_client à medida que são criados em _bootstrap_backend,
    para que o cleanup em main() consiga liberá-los mesmo se o bootstrap falhar a meio.
    """

    bundle: Any | None = None
    api_client: Optional[SomaApiClient] = None


@dataclass
class RunTotals:
    processed: int = 0
    ok: int = 0
    err: int = 0
    created: int = 0
    recovered: int = 0
    duplicated: int = 0
    transfer: int = 0
    row_times_ms: list[int] = field(default_factory=list)


@dataclass
class RowOutcome:
    ok: bool
    elapsed_ms: int
    created: bool = False
    recovered: bool = False
    duplicated: bool = False
    transfer: bool = False


def _bootstrap_backend(
    state: RunState,
    settings: Any,
    *,
    run_id: str,
    ws: str,
    headless: bool,
    backend_mode: str,
    api_first: bool,
    run_caixas_bancos: bool,
    api_fallback_selenium: bool,
) -> None:
    with step(
        logger,
        "run.init",
        run_id=run_id,
        sheet=ws,
        headless=headless,
        backend=backend_mode,
        caixas_bancos=run_caixas_bancos,
        api_fallback=api_fallback_selenium,
    ):
        needs_browser = (not api_first) or run_caixas_bancos or (api_first and api_fallback_selenium)
        if needs_browser:
            logger.warning("Validando a versão do ChromeDriver (inicialização do browser)...")
            t0_drv = time.perf_counter()

            state.bundle = WebDriverFactory.create(settings)

            dt_drv_ms = int((time.perf_counter() - t0_drv) * 1000)
            driver_obj = getattr(state.bundle, "a", None)
            wd = unwrap_webdriver(driver_obj)

            # =================================================================================
            # === A CURA PARA A CEGUEIRA NO SERVIDOR LINUX (FORÇAR RESOLUÇÃO FULL HD) ===
            # =================================================================================
            if wd:
                try:
                    wd.set_window_size(1920, 1080)
                except Exception as e:
                    logger.debug("Falha ao forçar window_size: %s", e)
            # =================================================================================

            info = get_chromedriver_info(wd)
            chrome_ver = get_chrome_version(wd)

            logger.warning(
                "ChromeDriver validado | chromedriver=%s | chrome=%s | path=%s | source=%s | dt_ms=%s",
                (info.get("version") or "desconhecida"),
                (chrome_ver or "desconhecida"),
                (info.get("path") or "n/a"),
                (info.get("source") or "unknown"),
                dt_drv_ms,
            )

        if api_first:
            base_url = env_str("SOMA_API_BASE_URL", "https://api-production-082f6.up.railway.app").rstrip("/")
            api_login = env_str("SOMA_API_LOGIN", "")
            api_password = env_str("SOMA_API_PASSWORD", "")
            session_token = env_str("SOMA_SESSION_TOKEN", "")
            has_api_credentials = bool(api_login and api_password)
            api_auth_preference = env_str("SOMA_API_AUTH_PREFERENCE", "token_first")
            session_token_close = env_bool("SOMA_SESSION_TOKEN_CLOSE", False)

            state.api_client = SomaApiClient(
                base_url=base_url,
                login=api_login,
                password=api_password,
                timeout_seconds=int(env_str("SOMA_API_TIMEOUT", "40") or 40),
                max_retries=int(env_str("SOMA_API_RETRIES", "2") or 2),
                backoff_seconds=float(env_str("SOMA_API_BACKOFF", "0.9") or 0.9),
            )

            _bootstrap_api_session(
                state.api_client,
                session_token=session_token,
                has_api_credentials=has_api_credentials,
                auth_preference=api_auth_preference,
                close_on_exit=session_token_close,
            )

    if state.bundle is not None:
        with step(logger, "run.login_ui", run_id=run_id):
            login = LoginPage(state.bundle.a, settings)
            if hasattr(login, "run"):
                login.run()
            else:
                login.login()


def _build_processors(
    *,
    settings: Any,
    bundle: Any | None,
    api_client: Optional[SomaApiClient],
    api_first: bool,
    api_fallback_selenium: bool,
) -> tuple[Any, Any]:
    if api_first:
        if api_client is None:
            raise RuntimeError("API client não inicializado (modo SOMA_BACKEND=api).")

        entradas_api = EntradasSaidasApi(api_client)
        transfer_api = TransferenciasApi(api_client)

        if api_fallback_selenium and bundle is not None:
            entradas_ui = EntradasSaidasPage(bundle.a, settings)
            transfer_ui = TransferenciasPage(bundle.a, settings)
            return (
                EntradasSaidasAdapter(entradas_api, entradas_ui),
                TransferenciasAdapter(transfer_api, transfer_ui),
            )
        return EntradasSaidasAdapter(entradas_api, None), TransferenciasAdapter(transfer_api, None)

    if bundle is None:
        raise RuntimeError("Browser não inicializado em modo Selenium.")
    return EntradasSaidasPage(bundle.a, settings), TransferenciasPage(bundle.a, settings)


def _process_row(
    table: SheetsTable,
    raw_row: Dict[str, Any],
    *,
    run_id: str,
    batch: int,
    progress_current: int,
    progress_total: int,
    iduser: str,
    allow_retry: bool,
    entradas_saidas: Any,
    transferencias: Any,
) -> RowOutcome:
    row_idx = int(raw_row["row"])
    tipo_txt = str(raw_row.get("TIPO") or raw_row.get("tipo") or "").strip() or "-"

    os.environ["ROW_NUMBER"] = str(row_idx)
    row_t0 = time.perf_counter()

    try:
        with step(
            logger,
            "run.process_row",
            run_id=run_id,
            batch=batch,
            row=row_idx,
            tipo=tipo_txt,
            progress_current=progress_current,
            progress_total=progress_total,
        ):
            row_model = ContaOrdemRow.from_table_row(row_number=row_idx, raw=raw_row)

            if row_model.tipo == TipoMovimento.TRANSFERENCIA:
                doc_id = transferencias.run(row_model)
                elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
                mark_row_ok(table, row_idx, str(doc_id), iduser, elapsed_ms=elapsed_ms)
                return RowOutcome(ok=True, elapsed_ms=elapsed_ms, transfer=True)

            duplicate_doc = entradas_saidas.precheck_duplicate(row_model)
            if duplicate_doc:
                elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
                mark_row_duplicate(table, row_idx, iduser, elapsed_ms=elapsed_ms)
                log_kv(
                    logger,
                    "Documento já existente. Lançamento será pulado.",
                    level=logging.INFO,
                    row=row_idx,
                    tipo=tipo_txt,
                    doc=duplicate_doc,
                )
                return RowOutcome(ok=True, elapsed_ms=elapsed_ms, duplicated=True)

            doc_id = entradas_saidas.create_and_get_doc_id(row_model)

            if not str(doc_id).startswith("SEM_DOC_"):
                mark_row_doc_soma(table, row_idx, str(doc_id))
                log_kv(
                    logger,
                    "DOC. SOMA gravado antecipadamente na sheet.",
                    level=logging.INFO,
                    row=row_idx,
                    tipo=tipo_txt,
                    doc=doc_id,
                )

            dados_doc = None
            if str(doc_id).startswith("SEM_DOC_"):
                dados_doc = None
            else:
                try:
                    dados_doc = entradas_saidas.fetch_dados_doc(str(doc_id))
                except Exception:
                    dados_doc = None

                try:
                    go_back = getattr(entradas_saidas, "go_back_to_list", None)
                    if callable(go_back):
                        go_back(row_model)
                except Exception:
                    logger.exception("Falha ao tentar voltar para a lista após capturar o DOC.")

            elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
            mark_row_ok(
                table,
                row_idx,
                str(doc_id),
                iduser,
                dados_doc=dados_doc,
                elapsed_ms=elapsed_ms,
            )
            return RowOutcome(
                ok=True,
                elapsed_ms=elapsed_ms,
                created=True,
                recovered=False,
            )

    except Exception as e:
        # Fora do `with step(...)`: deixa o step ver a exceção e emitir
        # [STEP_FAIL] com o traceback real, em vez de sempre [STEP_OK].
        elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
        err_msg = _safe_err(e)

        if _is_pending_doc_exception(e):
            mark_row_error(
                table,
                row_idx,
                err_msg,
                allow_retry=False,
                force_doc="EM ERRO",
                force_status="PENDENTE_DOC",
                elapsed_ms=elapsed_ms,
            )
        else:
            mark_row_error(
                table,
                row_idx,
                err_msg,
                allow_retry=allow_retry,
                elapsed_ms=elapsed_ms,
            )

        return RowOutcome(ok=False, elapsed_ms=elapsed_ms)


def _run_batches(
    *,
    settings: Any,
    ws: str,
    run_id: str,
    sheets: SheetsClient,
    result: Any,
    batch: int,
    bundle: Any | None,
    iduser: str,
    allow_retry: bool,
    run_caixas_bancos: bool,
    debug_row: int | None = None,
    entradas_saidas: Any,
    transferencias: Any,
) -> RunTotals:
    totals = RunTotals()
    attempted_rows: set[int] = set()

    while True:
        processable_rows: list[Dict[str, Any]] = []
        for raw_row in result.workset:
            row_number = int(raw_row["row"])
            if row_number in attempted_rows:
                continue
            processable_rows.append(raw_row)

        if not processable_rows:
            logger.warning(
                "Sem novas linhas para processar nesta execução. "
                "As restantes já foram tentadas no mesmo run_id."
            )
            break

        table = SheetsTable(sheets, ws)
        table.load()
        run_rows: Dict[int, Dict[str, Any]] = {int(r["row"]): r for r in processable_rows}

        try:
            total_processable = len(processable_rows)
            for progress_current, r in enumerate(processable_rows, start=1):
                row_idx = int(r["row"])
                attempted_rows.add(row_idx)

                outcome = _process_row(
                    table,
                    r,
                    run_id=run_id,
                    batch=batch,
                    progress_current=progress_current,
                    progress_total=total_processable,
                    iduser=iduser,
                    allow_retry=allow_retry,
                    entradas_saidas=entradas_saidas,
                    transferencias=transferencias,
                )

                totals.row_times_ms.append(outcome.elapsed_ms)
                totals.processed += 1
                if outcome.ok:
                    totals.ok += 1
                    if outcome.transfer:
                        totals.transfer += 1
                    elif outcome.duplicated:
                        totals.duplicated += 1
                    elif outcome.recovered:
                        totals.recovered += 1
                    else:
                        totals.created += 1
                else:
                    totals.err += 1

        finally:
            if allow_retry:
                table.load()
                current_rows = {int(x["row"]): x for x in table.get_records_with_row()}
                run_rows_live = {idx: current_rows.get(idx, {}) for idx in run_rows.keys()}
                unlock_still_processing(table, run_rows_live)

        batch += 1
        with step(logger, "run.preprocess", run_id=run_id, batch=batch):
            sheets = SheetsClient(settings)
            result = preprocess_contaordem(sheets, ws=ws, run_id=run_id, batch=batch, debug_row=debug_row)
        legacy_report.preprocess_summary(result.eligible_total, len(result.workset))

        if not result.workset:
            _run_post_processes(
                settings=settings,
                bundle=bundle,
                run_caixas_bancos=run_caixas_bancos,
            )
            break

    return totals


def _build_summary(
    *,
    run_id: str,
    ws: str,
    backend_mode: str,
    totals: RunTotals,
    total_elapsed_ms: int,
) -> str:
    avg_ms = int(sum(totals.row_times_ms) / len(totals.row_times_ms)) if totals.row_times_ms else 0
    return (
        "\n==================================================\n"
        "RESUMO FINAL\n"
        f"run_id: {run_id}\n"
        f"sheet: {ws}\n"
        f"backend: {backend_mode}\n"
        f"processadas: {totals.processed}\n"
        f"ok: {totals.ok}\n"
        f"erro: {totals.err}\n"
        f"criadas: {totals.created}\n"
        f"recuperadas: {totals.recovered}\n"
        f"duplicadas: {totals.duplicated}\n"
        f"transferencias: {totals.transfer}\n"
        f"tempo_total: {total_elapsed_ms/1000.0:.2f}s ({total_elapsed_ms}ms)\n"
        f"tempo_medio_linha: {avg_ms/1000.0:.2f}s ({avg_ms}ms)\n"
        "==================================================\n"
    )


def main() -> int:
    overall_t0 = time.perf_counter()

    settings = _load_settings()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    ws = _sheet_name(settings)
    run_id = new_run_id()
    os.environ["RUN_ID"] = run_id

    headless = env_bool("HEADLESS", True)
    allow_retry = env_bool("ALLOW_RETRY_ERROR", False)
    iduser = (env_str("IDUSER", "USERJOB") or "USERJOB").strip() or "USERJOB"
    keep_browser_open = env_bool("KEEP_BROWSER_OPEN", default=env_bool("DEBUG_STEP_MODE", False))

    backend_mode = env_str("SOMA_BACKEND", "selenium").lower()
    api_first = backend_mode == "api"

    run_caixas_bancos = env_bool("RUN_CAIXAS_BANCOS", default=True)
    api_fallback_selenium = env_bool("API_FALLBACK_SELENIUM", default=True)
    debug_row_raw = (env_str("DEBUG_ROW", "") or "").strip()
    debug_row = int(debug_row_raw) if debug_row_raw.isdigit() and int(debug_row_raw) > 0 else None

    state = RunState()
    totals = RunTotals()

    batch = 1
    with step(logger, "run.preprocess", run_id=run_id, batch=batch, initial=True):
        sheets = SheetsClient(settings)
        result = preprocess_contaordem(sheets, ws=ws, run_id=run_id, batch=batch, debug_row=debug_row)
    legacy_report.preprocess_summary(result.eligible_total, len(result.workset))

    if not result.workset:
        logger.warning("Nenhum documento pendente. Execução encerrada sem login/acesso web.")
    else:
        try:
            _bootstrap_backend(
                state,
                settings,
                run_id=run_id,
                ws=ws,
                headless=headless,
                backend_mode=backend_mode,
                api_first=api_first,
                run_caixas_bancos=run_caixas_bancos,
                api_fallback_selenium=api_fallback_selenium,
            )

            entradas_saidas, transferencias = _build_processors(
                settings=settings,
                bundle=state.bundle,
                api_client=state.api_client,
                api_first=api_first,
                api_fallback_selenium=api_fallback_selenium,
            )

            totals = _run_batches(
                settings=settings,
                ws=ws,
                run_id=run_id,
                sheets=sheets,
                result=result,
                batch=batch,
                bundle=state.bundle,
                iduser=iduser,
                allow_retry=allow_retry,
                run_caixas_bancos=run_caixas_bancos,
                debug_row=debug_row,
                entradas_saidas=entradas_saidas,
                transferencias=transferencias,
            )

        finally:
            try:
                if state.api_client is not None:
                    state.api_client.close_session()
            except Exception:
                pass
            try:
                if state.bundle is not None and not keep_browser_open:
                    state.bundle.quit()
                elif state.bundle is not None:
                    logger.warning(
                        "Browser mantido aberto ao final da execucao por KEEP_BROWSER_OPEN=true "
                        "ou DEBUG_STEP_MODE=true."
                    )
            except Exception:
                pass

    total_elapsed_ms = int((time.perf_counter() - overall_t0) * 1000)
    resumo_txt = _build_summary(
        run_id=run_id,
        ws=ws,
        backend_mode=backend_mode,
        totals=totals,
        total_elapsed_ms=total_elapsed_ms,
    )
    print(resumo_txt)
    logger.warning(resumo_txt.replace("\n", " | "))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
