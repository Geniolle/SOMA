from __future__ import annotations

import logging
import os
import re
import subprocess
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.sheets_client import SheetsClient
from soma_app.infra.trace import new_run_id, step
from soma_app.infra.webdriver_factory import WebDriverFactory
from soma_app.workflows.process_contaordem import (
    DOC_COL_DEFAULT,
    LOCK_VALUE_DEFAULT,
    STATUS_COL_DEFAULT,
    SheetsTable,
    preprocess_contaordem,
)
from soma_app.workflows.process_caixas_bancos import atualizar_caixas_bancos
from soma_app.workflows.process_soma import atualizar_sheet_soma

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.automation.pages.login_page import LoginPage
from soma_app.automation.pages.transferencias_page import TransferenciasPage

from soma_app.infra.soma_api_client import SomaApiClient
from soma_app.automation.api.iv_api import EntradasSaidasApi, TransferenciasApi

logger = logging.getLogger(__name__)

try:
    from dotenv import load_dotenv  # type: ignore

    load_dotenv()
except Exception:
    pass


def _bool_env(name: str, default: bool = False) -> bool:
    v = (os.getenv(name) or "").strip().lower()
    if v in {"1", "true", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "no", "n", "off"}:
        return False
    return default


def _env_str(name: str, default: str = "") -> str:
    v = (os.getenv(name) or "").strip()
    return v if v else default


def _is_pending_doc_exception(e: Exception) -> bool:
    try:
        tb = traceback.extract_tb(e.__traceback__)
        fnames = {f.name for f in tb}
        return "_go_back_to_list_best_effort" in fnames or "_search_doc_id" in fnames
    except Exception:
        return False


def _now_pt() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def _safe_err(e: Exception) -> str:
    s = str(e).strip().replace("\n", " ")
    return s[:180] if s else type(e).__name__


def _time_col_name(t: SheetsTable) -> Optional[str]:
    for c in ("TEMPO_PROCESSAMENTO_MS", "TEMPO_LINHA_MS", "PROCESS_TIME_MS", "DURACAO_MS"):
        if t.has_col(c):
            return c
    return None


# -----------------------------------------------------------------------------
# ChromeDriver version logging helpers
# -----------------------------------------------------------------------------
def _unwrap_webdriver(obj: Any) -> Any:
    """
    bundle.a pode ser um wrapper. Tenta chegar ao webdriver real.
    """
    cur = obj
    seen = set()
    for _ in range(6):
        if cur is None:
            return None
        oid = id(cur)
        if oid in seen:
            return cur
        seen.add(oid)

        if hasattr(cur, "capabilities") and (hasattr(cur, "execute_script") or hasattr(cur, "execute_cdp_cmd")):
            return cur

        for attr in ("driver", "_driver", "webdriver", "_webdriver", "wd", "_wd", "browser", "_browser"):
            nxt = getattr(cur, attr, None)
            if nxt is not None and nxt is not cur:
                cur = nxt
                break
        else:
            return cur
    return cur


def _get_driver_path(driver: Any) -> str:
    d = _unwrap_webdriver(driver)
    try:
        svc = getattr(d, "service", None) or getattr(d, "_service", None)
        p = getattr(svc, "path", None)
        if isinstance(p, str) and p.strip():
            return p.strip()
    except Exception:
        pass
    return ""


def _chromedriver_version_from_exe(driver_path: str) -> str:
    p = (driver_path or "").strip()
    if not p:
        return ""
    try:
        r = subprocess.run([p, "--version"], capture_output=True, text=True, timeout=5)
        out = (r.stdout or r.stderr or "").strip()
        m = re.search(r"ChromeDriver\s+(\d+\.\d+\.\d+\.\d+)", out)
        return m.group(1) if m else out[:120]
    except Exception:
        return ""


def _chromedriver_version_from_capabilities(driver: Any) -> str:
    d = _unwrap_webdriver(driver)
    try:
        caps = getattr(d, "capabilities", {}) or {}
        if isinstance(caps, dict):
            chrome = caps.get("chrome")
            if isinstance(chrome, dict):
                v = chrome.get("chromedriverVersion")
                if isinstance(v, str) and v.strip():
                    return v.split(" ")[0].strip()
    except Exception:
        pass
    return ""


def _find_chromedriver_in_known_caches() -> str:
    """
    Fallback: procura chromedriver.exe nos caches comuns:
      - Selenium Manager: %USERPROFILE%\.cache\selenium\...
      - webdriver_manager: %USERPROFILE%\.wdm\...
    """
    candidates: list[Path] = []

    home = Path.home()
    roots = [
        home / ".cache" / "selenium",
        home / ".wdm",
    ]

    la = os.getenv("LOCALAPPDATA")
    if la:
        roots.append(Path(la) / "selenium")
    tmp = os.getenv("TEMP")
    if tmp:
        roots.append(Path(tmp) / "selenium")

    for root in roots:
        try:
            if not root.exists():
                continue
            for p in root.rglob("chromedriver.exe"):
                candidates.append(p)
        except Exception:
            continue

    if not candidates:
        return ""

    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0])


def _get_chrome_version(driver: Any) -> str:
    d = _unwrap_webdriver(driver)

    try:
        caps = getattr(d, "capabilities", {}) or {}
        if isinstance(caps, dict):
            v = caps.get("browserVersion")
            if isinstance(v, str) and v.strip():
                return v.strip()
    except Exception:
        pass

    try:
        if hasattr(d, "execute_cdp_cmd"):
            info = d.execute_cdp_cmd("Browser.getVersion", {})
            if isinstance(info, dict):
                prod = info.get("product")
                if isinstance(prod, str) and prod.strip():
                    return prod.strip()
    except Exception:
        pass

    return ""


def _get_chromedriver_info(driver: Any) -> Dict[str, str]:
    d = _unwrap_webdriver(driver)

    v = _chromedriver_version_from_capabilities(d)
    if v:
        return {"version": v, "path": _get_driver_path(d) or "n/a", "source": "capabilities"}

    p = _get_driver_path(d)
    if p:
        v2 = _chromedriver_version_from_exe(p)
        if v2:
            return {"version": v2, "path": p, "source": "service.path"}

    p3 = _find_chromedriver_in_known_caches()
    if p3:
        v3 = _chromedriver_version_from_exe(p3)
        if v3:
            return {"version": v3, "path": p3, "source": "cache"}

    return {"version": "", "path": p or "n/a", "source": "unknown"}


def _mark_row_ok(
    t: SheetsTable,
    row_idx: int,
    doc_id: str,
    iduser: str,
    dados_doc: str | None = None,
    elapsed_ms: int | None = None,
) -> None:
    updates = [
        (row_idx, DOC_COL_DEFAULT, doc_id),
        (row_idx, "IDUSER", iduser),
        (row_idx, "TIMESTAMP", _now_pt()),
    ]
    if t.has_col(STATUS_COL_DEFAULT):
        updates.append((row_idx, STATUS_COL_DEFAULT, "VALIDADO"))
    if dados_doc and t.has_col("DADOS DOC"):
        updates.append((row_idx, "DADOS DOC", dados_doc))
    if elapsed_ms is not None:
        c = _time_col_name(t)
        if c:
            updates.append((row_idx, c, elapsed_ms))
    t.batch_update_cells(updates)


def _mark_row_error(
    t: SheetsTable,
    row_idx: int,
    err_msg: str,
    allow_retry: bool,
    *,
    force_doc: str | None = None,
    force_status: str | None = None,
    elapsed_ms: int | None = None,
) -> None:
    doc_value = force_doc if force_doc is not None else ("" if allow_retry else "EM ERRO")
    updates = [(row_idx, DOC_COL_DEFAULT, doc_value)]

    if t.has_col(STATUS_COL_DEFAULT):
        status_value = force_status if force_status is not None else ("VALIDADO" if allow_retry else "ERRO")
        updates.append((row_idx, STATUS_COL_DEFAULT, status_value))

    for col in ("ULTIMO_ERRO", "ÚLTIMO_ERRO", "ERRO", "ERROR_MSG"):
        if t.has_col(col):
            updates.append((row_idx, col, err_msg))
            break

    if t.has_col("TIMESTAMP"):
        updates.append((row_idx, "TIMESTAMP", _now_pt()))

    if elapsed_ms is not None:
        c = _time_col_name(t)
        if c:
            updates.append((row_idx, c, elapsed_ms))

    t.batch_update_cells(updates)


def _unlock_still_processing(t: SheetsTable, run_rows: Dict[int, Dict[str, Any]]) -> None:
    updates = []
    for row_idx, r in run_rows.items():
        doc = str(r.get(DOC_COL_DEFAULT, "")).strip()
        if doc.lower() == LOCK_VALUE_DEFAULT.lower():
            updates.append((row_idx, DOC_COL_DEFAULT, ""))
            if t.has_col(STATUS_COL_DEFAULT):
                updates.append((row_idx, STATUS_COL_DEFAULT, "VALIDADO"))
    if updates:
        t.batch_update_cells(updates)


class _EntradasSaidasAdapter:
    def __init__(self, primary: Any, fallback: Any | None = None):
        self.primary = primary
        self.fallback = fallback

    def create_and_get_doc_id(self, row: ContaOrdemRow) -> str:
        try:
            return self.primary.create_and_get_doc_id(row)
        except Exception as e:
            if self.fallback is None:
                raise
            logger.warning("API falhou (create Entradas/Saídas). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.create_and_get_doc_id(row)

    def recover_doc_id(self, row: ContaOrdemRow) -> str:
        try:
            return self.primary.recover_doc_id(row)
        except Exception as e:
            if self.fallback is None:
                raise
            logger.warning("API falhou (recover DOC). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.recover_doc_id(row)

    def fetch_dados_doc(self, doc_id: str) -> str:
        return self.primary.fetch_dados_doc(doc_id)


class _TransferenciasAdapter:
    def __init__(self, primary: Any, fallback: Any | None = None):
        self.primary = primary
        self.fallback = fallback

    def run(self, row: ContaOrdemRow) -> str:
        try:
            return self.primary.run(row)
        except Exception as e:
            if self.fallback is None:
                raise
            logger.warning("API falhou (Transferência). Vou tentar Selenium. err=%s", _safe_err(e))
            return self.fallback.run(row)


def _load_settings() -> Any:
    from soma_app.config.settings import Settings

    env_file = os.getenv("ENV_FILE")
    return Settings.from_env(env_file=env_file if env_file else None)


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


def main() -> int:
    overall_t0 = time.perf_counter()

    settings = _load_settings()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    ws = _sheet_name(settings)
    run_id = new_run_id()
    os.environ["RUN_ID"] = run_id

    headless = _bool_env("HEADLESS", True)
    allow_retry = _bool_env("ALLOW_RETRY_ERROR", False)
    iduser = (_env_str("IDUSER", "USERJOB") or "USERJOB").strip() or "USERJOB"

    backend_mode = _env_str("SOMA_BACKEND", "selenium").lower()
    api_first = backend_mode == "api"

    run_caixas_bancos = _bool_env("RUN_CAIXAS_BANCOS", default=True)
    api_fallback_selenium = _bool_env("API_FALLBACK_SELENIUM", default=True)

    bundle = None
    api_client: Optional[SomaApiClient] = None

    total_rows_processed = 0
    total_ok = 0
    total_err = 0
    total_created = 0
    total_recovered = 0
    total_transfer = 0
    row_times_ms: list[int] = []

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
        sheets = SheetsClient(settings)

        needs_browser = (not api_first) or run_caixas_bancos or (api_first and api_fallback_selenium)
        if needs_browser:
            logger.warning("Validando a versão do ChromeDriver (inicialização do browser)...")
            t0_drv = time.perf_counter()

            bundle = WebDriverFactory.create(settings)

            dt_drv_ms = int((time.perf_counter() - t0_drv) * 1000)
            driver_obj = getattr(bundle, "a", None)
            wd = _unwrap_webdriver(driver_obj)

            info = _get_chromedriver_info(wd)
            chrome_ver = _get_chrome_version(wd)

            logger.warning(
                "ChromeDriver validado | chromedriver=%s | chrome=%s | path=%s | source=%s | dt_ms=%s",
                (info.get("version") or "desconhecida"),
                (chrome_ver or "desconhecida"),
                (info.get("path") or "n/a"),
                (info.get("source") or "unknown"),
                dt_drv_ms,
            )

        if api_first:
            base_url = _env_str("SOMA_API_BASE_URL", "https://api-production-082f6.up.railway.app").rstrip("/")

            api_client = SomaApiClient(
                base_url=base_url,
                login=_env_str("SOMA_API_LOGIN", ""),
                password=_env_str("SOMA_API_PASSWORD", ""),
                timeout_seconds=int(_env_str("SOMA_API_TIMEOUT", "40") or 40),
                max_retries=int(_env_str("SOMA_API_RETRIES", "2") or 2),
                backoff_seconds=float(_env_str("SOMA_API_BACKOFF", "0.9") or 0.9),
            )

            sess = _env_str("SOMA_SESSION_TOKEN", "")
            if sess:
                api_client.use_session_token(sess, close_on_exit=_bool_env("SOMA_SESSION_TOKEN_CLOSE", False))
            else:
                api_client.open_session()

    if bundle is not None:
        with step(logger, "run.login_ui", run_id=run_id):
            login = LoginPage(bundle.a, settings)
            if hasattr(login, "run"):
                login.run()
            else:
                login.login()

    entradas_saidas: Any
    transferencias: Any

    if api_first:
        if api_client is None:
            raise RuntimeError("API client não inicializado (modo SOMA_BACKEND=api).")

        entradas_api = EntradasSaidasApi(api_client)
        transfer_api = TransferenciasApi(api_client)

        if api_fallback_selenium and bundle is not None:
            entradas_ui = EntradasSaidasPage(bundle.a, settings)
            transfer_ui = TransferenciasPage(bundle.a, settings)
            entradas_saidas = _EntradasSaidasAdapter(entradas_api, entradas_ui)
            transferencias = _TransferenciasAdapter(transfer_api, transfer_ui)
        else:
            entradas_saidas = _EntradasSaidasAdapter(entradas_api, None)
            transferencias = _TransferenciasAdapter(transfer_api, None)
    else:
        if bundle is None:
            raise RuntimeError("Browser não inicializado em modo Selenium.")
        entradas_saidas = EntradasSaidasPage(bundle.a, settings)
        transferencias = TransferenciasPage(bundle.a, settings)

    batch = 1
    try:
        while True:
            with step(logger, "run.preprocess", run_id=run_id, batch=batch):
                sheets = SheetsClient(settings)
                result = preprocess_contaordem(sheets, ws=ws, run_id=run_id, batch=batch)

            if not result.workset:
                _run_post_processes(
                    settings=settings,
                    bundle=bundle,
                    run_caixas_bancos=run_caixas_bancos,
                )
                break

            table = SheetsTable(sheets, ws)
            table.load()
            run_rows: Dict[int, Dict[str, Any]] = {int(r["row"]): r for r in result.workset}

            try:
                for r in result.workset:
                    row_idx = int(r["row"])
                    tipo_txt = str(r.get("TIPO") or r.get("tipo") or "").strip() or "-"

                    doc_cell = str(r.get(DOC_COL_DEFAULT, "") or "").strip().upper()
                    status_cell = str(r.get(STATUS_COL_DEFAULT, "") or "").strip().upper()
                    is_recover = doc_cell == "EM ERRO" and status_cell == "PENDENTE_DOC"

                    os.environ["ROW_NUMBER"] = str(row_idx)
                    row_t0 = time.perf_counter()

                    with step(logger, "run.process_row", run_id=run_id, batch=batch, row=row_idx, tipo=tipo_txt):
                        try:
                            row_model = ContaOrdemRow.from_table_row(row_number=row_idx, raw=r)

                            if row_model.tipo == TipoMovimento.TRANSFERENCIA:
                                doc_id = transferencias.run(row_model)
                                elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
                                row_times_ms.append(elapsed_ms)
                                _mark_row_ok(table, row_idx, str(doc_id), iduser, elapsed_ms=elapsed_ms)

                                total_rows_processed += 1
                                total_ok += 1
                                total_transfer += 1
                                continue

                            if is_recover:
                                doc_id = entradas_saidas.recover_doc_id(row_model)
                                total_recovered += 1
                            else:
                                doc_id = entradas_saidas.create_and_get_doc_id(row_model)
                                total_created += 1

                            dados_doc = None
                            try:
                                dados_doc = entradas_saidas.fetch_dados_doc(str(doc_id))
                            except Exception:
                                dados_doc = None

                            elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
                            row_times_ms.append(elapsed_ms)
                            _mark_row_ok(
                                table,
                                row_idx,
                                str(doc_id),
                                iduser,
                                dados_doc=dados_doc,
                                elapsed_ms=elapsed_ms,
                            )

                            total_rows_processed += 1
                            total_ok += 1

                        except Exception as e:
                            elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
                            row_times_ms.append(elapsed_ms)
                            err_msg = _safe_err(e)

                            if _is_pending_doc_exception(e):
                                _mark_row_error(
                                    table,
                                    row_idx,
                                    err_msg,
                                    allow_retry=False,
                                    force_doc="EM ERRO",
                                    force_status="PENDENTE_DOC",
                                    elapsed_ms=elapsed_ms,
                                )
                            else:
                                _mark_row_error(table, row_idx, err_msg, allow_retry=allow_retry, elapsed_ms=elapsed_ms)

                            total_rows_processed += 1
                            total_err += 1

            finally:
                if allow_retry:
                    table.load()
                    current_rows = {int(x["row"]): x for x in table.get_records_with_row()}
                    run_rows_live = {idx: current_rows.get(idx, {}) for idx in run_rows.keys()}
                    _unlock_still_processing(table, run_rows_live)

            batch += 1

    finally:
        try:
            if api_client is not None:
                api_client.close_session()
        except Exception:
            pass
        try:
            if bundle is not None:
                bundle.quit()
        except Exception:
            pass

    total_elapsed_ms = int((time.perf_counter() - overall_t0) * 1000)
    avg_ms = int(sum(row_times_ms) / len(row_times_ms)) if row_times_ms else 0

    resumo_txt = (
        "\n==================================================\n"
        "RESUMO FINAL\n"
        f"run_id: {run_id}\n"
        f"sheet: {ws}\n"
        f"backend: {backend_mode}\n"
        f"processadas: {total_rows_processed}\n"
        f"ok: {total_ok}\n"
        f"erro: {total_err}\n"
        f"criadas: {total_created}\n"
        f"recuperadas: {total_recovered}\n"
        f"transferencias: {total_transfer}\n"
        f"tempo_total: {total_elapsed_ms/1000.0:.2f}s ({total_elapsed_ms}ms)\n"
        f"tempo_medio_linha: {avg_ms/1000.0:.2f}s ({avg_ms}ms)\n"
        "==================================================\n"
    )
    print(resumo_txt)
    logger.warning(resumo_txt.replace("\n", " | "))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
