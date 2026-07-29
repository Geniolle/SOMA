from __future__ import annotations

import logging
import os
import re
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.automation.pages.login_page import LoginPage
from soma_app.automation.pages.transferencias_page import TransferenciasPage
from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra import report as legacy_report
from soma_app.infra.audit import audit_event, audit_row
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
from soma_app.workflows.contaordem_writer import mark_row_error, mark_row_ok, unlock_still_processing
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


def _progress(enabled: bool, msg: str, **fields: Any) -> None:
    if not enabled:
        return
    log_kv(logger, msg, level=logging.INFO, **fields)


def _norm_status(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").strip().upper())


def _should_recover_doc(raw_row: Dict[str, Any]) -> bool:
    # STATUS=PENDENTE_DOC Ã© o Ãºnico sinal confiÃ¡vel: DOC. SOMA pode ter sido
    # sobrescrito para "Em processamento" (lock) ou "" (unlock_still_processing)
    # entre o run que criou o documento e este reprocessamento.
    status_value = raw_row.get(STATUS_COL_DEFAULT, "")
    return _norm_status(status_value) == "PENDENTEDOC"


def _bootstrap_api_session(
    api_client: Any,
    *,
    session_token: str = "",
    has_api_credentials: bool = False,
    auth_preference: str = "token_first",
    close_on_exit: bool = False,
) -> None:
    token = (session_token or "").strip()
    preference = (auth_preference or "token_first").strip().lower()

    if preference == "token_first":
        if token:
            api_client.use_session_token(token, close_on_exit=close_on_exit)
            return
        if has_api_credentials:
            api_client.open_session()
        return

    if preference == "login_first":
        if has_api_credentials:
            try:
                api_client.open_session()
                return
            except Exception:
                if token:
                    logger.warning("Usando SOMA_SESSION_TOKEN como fallback.")
                    api_client.use_session_token(token, close_on_exit=close_on_exit)
                    return
                raise
        if token:
            api_client.use_session_token(token, close_on_exit=close_on_exit)
        return

    if token:
        api_client.use_session_token(token, close_on_exit=close_on_exit)
    elif has_api_credentials:
        api_client.open_session()


def _load_settings() -> Any:
    from soma_app.config.settings import Settings

    # run_soma.py -> workflows -> soma_app -> src -> <project_root>
    project_root = Path(__file__).resolve().parents[3]

    env_file = (os.getenv("ENV_FILE") or "").strip()
    if env_file:
        env_path = Path(env_file)
        if not env_path.is_absolute():
            # 1) tenta relativo ao cwd (comportamento padrÃ£o)
            cwd_candidate = Path.cwd() / env_path
            if cwd_candidate.exists():
                return Settings.from_env(env_file=str(cwd_candidate))

            # 2) tenta relativo Ã  raiz do projeto
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
        logger.warning("PÃ³s-processos ignorados: browser indisponÃ­vel.")
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
        logger.warning("Processo de atualizaÃ§Ã£o da sheet SOMA nÃ£o serÃ¡ executado porque Caixas/Bancos falhou.")
        return

    with step(logger, "run.post.atualizar_sheet_soma"):
        try:
            sheets_soma = SheetsClient(settings)
            atualizar_sheet_soma(sheets_soma, bundle.a, settings)
        except Exception:
            logger.exception("Falha no processo de atualizaÃ§Ã£o da sheet SOMA.")


@dataclass
class RunState:
    """
    Acumula o bundle criado em _bootstrap_backend para que o cleanup em main()
    consiga liberÃ¡-lo mesmo se o bootstrap falhar a meio.
    """

    bundle: Any | None = None
    api_client: Any | None = None


@dataclass
class RunTotals:
    processed: int = 0
    ok: int = 0
    err: int = 0
    created: int = 0
    recovered: int = 0
    transfer: int = 0
    row_times_ms: list[int] = field(default_factory=list)


@dataclass
class RowOutcome:
    ok: bool
    elapsed_ms: int
    created: bool = False
    recovered: bool = False
    transfer: bool = False


def _bootstrap_backend(
    state: RunState,
    settings: Any,
    *,
    run_id: str,
    ws: str,
    headless: bool,
    run_caixas_bancos: bool,
    detailed_logging: bool = False,
    backend_mode: str = "selenium",
    api_first: bool = False,
    api_fallback_selenium: bool = True,
) -> None:
    _progress(
        detailed_logging,
        "A iniciar backend Selenium.",
        run_id=run_id,
        sheet=ws,
        headless=headless,
        caixas_bancos=run_caixas_bancos,
    )
    with step(
        logger,
        "run.init",
        run_id=run_id,
        sheet=ws,
        headless=headless,
        caixas_bancos=run_caixas_bancos,
    ):
        logger.warning("Validando a versÃ£o do ChromeDriver (inicializaÃ§Ã£o do browser)...")
        t0_drv = time.perf_counter()

        state.bundle = WebDriverFactory.create(settings)

        dt_drv_ms = int((time.perf_counter() - t0_drv) * 1000)
        driver_obj = getattr(state.bundle, "a", None)
        wd = unwrap_webdriver(driver_obj)

        # =================================================================================
        # === A CURA PARA A CEGUEIRA NO SERVIDOR LINUX (FORÃ‡AR RESOLUÃ‡ÃƒO FULL HD) ===
        # =================================================================================
        if wd:
            try:
                wd.set_window_size(1920, 1080)
            except Exception as e:
                logger.debug("Falha ao forÃ§ar window_size: %s", e)
        # =================================================================================

        info = get_chromedriver_info(wd)
        chrome_ver = get_chrome_version(wd)

        _progress(
            detailed_logging,
            "Backend Selenium pronto.",
            run_id=run_id,
            chromedriver=(info.get("version") or "desconhecida"),
            chrome=(chrome_ver or "desconhecida"),
            path=(info.get("path") or "n/a"),
            source=(info.get("source") or "unknown"),
            dt_ms=dt_drv_ms,
        )
        logger.warning(
            "ChromeDriver validado | chromedriver=%s | chrome=%s | path=%s | source=%s | dt_ms=%s",
            (info.get("version") or "desconhecida"),
            (chrome_ver or "desconhecida"),
            (info.get("path") or "n/a"),
            (info.get("source") or "unknown"),
            dt_drv_ms,
        )

    if backend_mode.strip().lower() == "api" and api_first:
        api_client = SomaApiClient(
            base_url=getattr(settings, "soma_api_base_url", "") or "",
            login=getattr(settings, "soma_api_login", "") or "",
            password=getattr(settings, "soma_api_password", "") or "",
            session_token=getattr(settings, "soma_session_token", "") or "",
            close_on_exit=bool(getattr(settings, "soma_session_token_close", False)),
            timeout_seconds=int(getattr(settings, "soma_api_timeout", 40) or 40),
            max_retries=int(getattr(settings, "soma_api_retries", 2) or 2),
            backoff_seconds=float(getattr(settings, "soma_api_backoff", 0.9) or 0.9),
        )
        state.api_client = api_client
        _bootstrap_api_session(
            api_client,
            session_token=getattr(settings, "soma_session_token", "") or "",
            has_api_credentials=bool((getattr(settings, "soma_api_login", "") or "").strip() and (getattr(settings, "soma_api_password", "") or "").strip()),
            auth_preference=(getattr(settings, "soma_api_auth_preference", "token_first") or "token_first"),
            close_on_exit=bool(getattr(settings, "soma_session_token_close", False)),
        )

    if state.bundle is not None:
        _progress(detailed_logging, "A iniciar login no portal SOMA.", run_id=run_id)
        with step(logger, "run.login_ui", run_id=run_id):
            login = LoginPage(state.bundle.a, settings)
            if hasattr(login, "run"):
                login.run()
            else:
                login.login()
        _progress(detailed_logging, "Login no portal SOMA concluído.", run_id=run_id)


def _build_processors(
    *,
    settings: Any,
    bundle: Any | None,
) -> tuple[Any, Any]:
    if bundle is None:
        raise RuntimeError("Browser nÃ£o inicializado em modo Selenium.")
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
    detailed_logging: bool = False,
    entradas_saidas: Any,
    transferencias: Any,
) -> RowOutcome:
    row_idx = int(raw_row["row"])
    tipo_txt = str(raw_row.get("TIPO") or raw_row.get("tipo") or "").strip() or "-"
    is_recover = _should_recover_doc(raw_row)

    os.environ["ROW_NUMBER"] = str(row_idx)
    row_t0 = time.perf_counter()
    _progress(
        detailed_logging,
        "A processar linha.",
        run_id=run_id,
        batch=batch,
        row=row_idx,
        tipo=tipo_txt,
        progress_current=progress_current,
        progress_total=progress_total,
        recover=is_recover,
    )

    try:
        with audit_row(
            run_id=run_id,
            batch=batch,
            row=row_idx,
            tipo=tipo_txt,
            payload={"progress_current": progress_current, "progress_total": progress_total},
        ):
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
                    _progress(
                        detailed_logging,
                        "Transferência concluída.",
                        run_id=run_id,
                        batch=batch,
                        row=row_idx,
                        doc_id=doc_id,
                        elapsed_ms=elapsed_ms,
                    )
                    audit_event(
                        "ROW_OK",
                        run_id=run_id,
                        batch=batch,
                        row=row_idx,
                        tipo=tipo_txt,
                        doc_id=doc_id,
                        result="transfer",
                    )
                    mark_row_ok(table, row_idx, str(doc_id), iduser, elapsed_ms=elapsed_ms)
                    return RowOutcome(ok=True, elapsed_ms=elapsed_ms, transfer=True)

                if is_recover:
                    try:
                        doc_id = entradas_saidas.recover_doc_id(row_model)
                    except Exception as recover_exc:
                        logger.warning(
                            "Recovery de DOC falhou; vou recriar o lancamento para manter o fluxo continuo. err=%s",
                            _safe_err(recover_exc),
                        )
                        _progress(
                            detailed_logging,
                            "Recovery de DOC falhou; a recriar lançamento.",
                            run_id=run_id,
                            batch=batch,
                            row=row_idx,
                            err=_safe_err(recover_exc),
                        )
                        doc_id = entradas_saidas.create_and_get_doc_id(row_model)
                else:
                    doc_id = entradas_saidas.create_and_get_doc_id(row_model)

                dados_doc = (row_model.dados_doc or "").strip() or None
                if not dados_doc:
                    try:
                        dados_doc = entradas_saidas.fetch_dados_doc(str(doc_id))
                    except Exception:
                        dados_doc = None

                elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
                _progress(
                    detailed_logging,
                    "Linha concluída.",
                    run_id=run_id,
                    batch=batch,
                    row=row_idx,
                    doc_id=doc_id,
                    recovered=is_recover,
                    created=not is_recover,
                    has_dados_doc=bool(dados_doc),
                    elapsed_ms=elapsed_ms,
                )
                audit_event(
                    "ROW_OK",
                    run_id=run_id,
                    batch=batch,
                    row=row_idx,
                    tipo=tipo_txt,
                    doc_id=doc_id,
                    recovered=is_recover,
                    created=not is_recover,
                    dados_doc=bool(dados_doc),
                )
                mark_row_ok(
                    table,
                    row_idx,
                    str(doc_id),
                    iduser,
                    dados_doc=dados_doc,
                    elapsed_ms=elapsed_ms,
                )
                return RowOutcome(ok=True, elapsed_ms=elapsed_ms, created=not is_recover, recovered=is_recover)

    except Exception as e:
        elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
        err_msg = _safe_err(e)
        _progress(
            detailed_logging,
            "Linha com erro.",
            run_id=run_id,
            batch=batch,
            row=row_idx,
            tipo=tipo_txt,
            err=err_msg,
            elapsed_ms=elapsed_ms,
        )

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
    detailed_logging: bool = False,
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
                "Sem novas linhas para processar nesta execuÃ§Ã£o. "
                "As restantes jÃ¡ foram tentadas no mesmo run_id."
            )
            break

        table = SheetsTable(sheets, ws)
        table.load()
        run_rows: Dict[int, Dict[str, Any]] = {int(r["row"]): r for r in processable_rows}

        try:
            total_processable = len(processable_rows)
            _progress(
                detailed_logging,
                "A iniciar lote.",
                run_id=run_id,
                batch=batch,
                rows=total_processable,
            )
            with step(logger, "run.batch", run_id=run_id, batch=batch, rows=total_processable):
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
                        detailed_logging=detailed_logging,
                        entradas_saidas=entradas_saidas,
                        transferencias=transferencias,
                    )

                    totals.row_times_ms.append(outcome.elapsed_ms)
                    totals.processed += 1
                    if outcome.ok:
                        totals.ok += 1
                        if outcome.transfer:
                            totals.transfer += 1
                        elif outcome.recovered:
                            totals.recovered += 1
                        else:
                            totals.created += 1
                    else:
                        totals.err += 1

            _progress(
                detailed_logging,
                "Lote concluído.",
                run_id=run_id,
                batch=batch,
                processed=totals.processed,
                ok=totals.ok,
                err=totals.err,
            )

        finally:
            if allow_retry:
                table.load()
                current_rows = {int(x["row"]): x for x in table.get_records_with_row()}
                run_rows_live = {idx: current_rows.get(idx, {}) for idx in run_rows.keys()}
                unlock_still_processing(table, run_rows_live)

        batch += 1
        _progress(detailed_logging, "A preparar próximo lote.", run_id=run_id, next_batch=batch)
        with step(logger, "run.preprocess", run_id=run_id, batch=batch):
            sheets = SheetsClient(settings)
            result = preprocess_contaordem(sheets, ws=ws, run_id=run_id, batch=batch)
        legacy_report.preprocess_summary(result.eligible_total, len(result.workset))

        if not result.workset:
            _progress(detailed_logging, "Fila esvaziada; a executar pós-processos.", run_id=run_id)
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
    totals: RunTotals,
    total_elapsed_ms: int,
) -> str:
    avg_ms = int(sum(totals.row_times_ms) / len(totals.row_times_ms)) if totals.row_times_ms else 0
    return (
        "\n==================================================\n"
        "RESUMO FINAL\n"
        f"run_id: {run_id}\n"
        f"sheet: {ws}\n"
        "backend: selenium\n"
        f"processadas: {totals.processed}\n"
        f"ok: {totals.ok}\n"
        f"erro: {totals.err}\n"
        f"criadas: {totals.created}\n"
        f"recuperadas: {totals.recovered}\n"
        f"transferencias: {totals.transfer}\n"
        f"tempo_total: {total_elapsed_ms/1000.0:.2f}s ({total_elapsed_ms}ms)\n"
        f"tempo_medio_linha: {avg_ms/1000.0:.2f}s ({avg_ms}ms)\n"
        "==================================================\n"
    )


def main() -> int:
    overall_t0 = time.perf_counter()

    settings = _load_settings()
    ws = _sheet_name(settings)
    run_id = new_run_id()
    os.environ["RUN_ID"] = run_id
    audit_event("RUN_START", run_id=run_id, sheet=ws)

    detailed_logging = env_bool("DETAILED_RUN_LOGGING", False)
    headless = env_bool("HEADLESS", True)
    allow_retry = env_bool("ALLOW_RETRY_ERROR", False)
    iduser = (env_str("IDUSER", "USERJOB") or "USERJOB").strip() or "USERJOB"

    run_caixas_bancos = env_bool("RUN_CAIXAS_BANCOS", default=True)

    configure_logging(settings, console_level="INFO" if detailed_logging else None)
    ensure_artifacts_dirs(settings)

    _progress(
        detailed_logging,
        "Execução iniciada.",
        run_id=run_id,
        sheet=ws,
        headless=headless,
        allow_retry=allow_retry,
        caixas_bancos=run_caixas_bancos,
    )

    state = RunState()
    totals = RunTotals()

    batch = 1
    with step(logger, "run.preprocess", run_id=run_id, batch=batch, initial=True):
        sheets = SheetsClient(settings)
        result = preprocess_contaordem(sheets, ws=ws, run_id=run_id, batch=batch)
    legacy_report.preprocess_summary(result.eligible_total, len(result.workset))
    _progress(
        detailed_logging,
        "Pré-processamento inicial concluído.",
        run_id=run_id,
        eligible=result.eligible_total,
        workset=len(result.workset),
    )

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
                run_caixas_bancos=run_caixas_bancos,
                detailed_logging=detailed_logging,
            )

            entradas_saidas, transferencias = _build_processors(
                settings=settings,
                bundle=state.bundle,
            )
            _progress(
                detailed_logging,
                "Processadores prontos.",
                run_id=run_id,
                processor_entradas_saidas=type(entradas_saidas).__name__,
                processor_transferencias=type(transferencias).__name__,
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
                detailed_logging=detailed_logging,
                entradas_saidas=entradas_saidas,
                transferencias=transferencias,
            )

        finally:
            try:
                if state.bundle is not None:
                    state.bundle.quit()
            except Exception:
                pass

    total_elapsed_ms = int((time.perf_counter() - overall_t0) * 1000)
    audit_event(
        "RUN_END",
        run_id=run_id,
        sheet=ws,
        backend="selenium",
        processed=totals.processed,
        ok=totals.ok,
        err=totals.err,
        created=totals.created,
        recovered=totals.recovered,
        transfer=totals.transfer,
        total_ms=total_elapsed_ms,
    )
    resumo_txt = _build_summary(
        run_id=run_id,
        ws=ws,
        totals=totals,
        total_elapsed_ms=total_elapsed_ms,
    )
    print(resumo_txt)
    logger.warning(resumo_txt.replace("\n", " | "))
    _progress(
        detailed_logging,
        "Execução concluída.",
        run_id=run_id,
        processed=totals.processed,
        ok=totals.ok,
        err=totals.err,
        created=totals.created,
        recovered=totals.recovered,
        transfer=totals.transfer,
        total_ms=total_elapsed_ms,
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())


