from __future__ import annotations

import logging
import os
import time
from dataclasses import dataclass
from typing import Any, Dict

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.automation.pages.login_page import LoginPage
from soma_app.domain.models import ContaOrdemRow, TipoMovimento, normalize_document_value
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.sheets_client import SheetsClient
from soma_app.infra.trace import log_kv, new_run_id, step
from soma_app.infra.webdriver_factory import WebDriverFactory
from soma_app.workflows.contaordem_writer import mark_row_auditoria
from soma_app.workflows.process_contaordem import SheetsTable
from soma_app.workflows.run_soma import _load_settings, _sheet_name

logger = logging.getLogger(__name__)


def _safe_err(e: Exception) -> str:
    s = str(e).strip().replace("\n", " ")
    return s[:180] if s else type(e).__name__


def _is_blank(value: Any) -> bool:
    return not str(value or "").strip()


_TIPOS_AUDITAVEIS = {TipoMovimento.ENTRADA, TipoMovimento.SAIDA}


def _is_entrada_ou_saida(raw: Dict[str, Any]) -> bool:
    """Retorna True apenas se TIPO for Entrada ou Saída (ignora Cartão, Transferência, etc.)."""
    try:
        tipo = TipoMovimento.from_sheet_value(raw.get("TIPO", ""))
        return tipo in _TIPOS_AUDITAVEIS
    except ValueError:
        return False


@dataclass
class AuditTotals:
    analyzed: int = 0
    confirmed: int = 0
    divergent: int = 0
    technical_errors: int = 0


@dataclass
class AuditOutcome:
    analyzed: bool
    confirmed: bool = False
    divergent: bool = False
    technical_error: bool = False


def _header() -> None:
    logger.warning("=" * 80)
    logger.warning("CONCILIAÇÃO SOMA")
    logger.warning("=" * 80)


def _log_row_context(row: ContaOrdemRow) -> None:
    logger.warning("")
    logger.warning("Linha %s", row.row_number)
    logger.warning("ID_INTERNO=%s", row.id_interno or "-")
    logger.warning("TIPO=%s", row.tipo.value)
    logger.warning("DATA=%s", row.data_mov or "-")
    logger.warning("DESCRICAO=%s", row.descricao_soma or "-")
    logger.warning("DOC_ESPERADO=%s", row.doc_soma or "-")


def _select_rows_to_audit(table: SheetsTable) -> list[Dict[str, Any]]:
    if not table.has_col("AUDITORIA"):
        raise RuntimeError("Coluna AUDITORIA nao encontrada na worksheet CONTAORDEM.")

    rows: list[Dict[str, Any]] = []
    skipped_tipo: list[str] = []
    for raw in table.get_records_with_row():
        if not _is_blank(raw.get("AUDITORIA")):
            continue
        if not _is_entrada_ou_saida(raw):
            skipped_tipo.append(str(raw.get("TIPO", "")).strip() or "(vazio)")
            continue
        rows.append(raw)

    if skipped_tipo:
        from collections import Counter
        counts = Counter(skipped_tipo)
        resumo = ", ".join(f"{t!r}: {n}" for t, n in counts.most_common())
        logger.info("Linhas ignoradas por TIPO não auditável: %s", resumo)

    return rows


def _compare_documents(expected_doc: Any, found_doc: Any) -> str:
    expected_norm = normalize_document_value(expected_doc)
    found_norm = normalize_document_value(found_doc)
    if not expected_norm:
        return "Divergente"
    return "Confirmado" if expected_norm == found_norm else "Divergente"


def _audit_row(
    *,
    table: SheetsTable,
    page: EntradasSaidasPage,
    raw_row: Dict[str, Any],
    run_id: str,
) -> AuditOutcome:
    row_idx = int(raw_row["row"])
    row_t0 = time.perf_counter()

    try:
        with step(logger, "audit.process_row", run_id=run_id, row=row_idx):
            row_model = ContaOrdemRow.from_table_row(row_number=row_idx, raw=raw_row)
            _log_row_context(row_model)
            logger.warning("Pesquisando documento no SOMA...")

            found_doc = page.search_existing_doc(row_model)
            expected_norm = normalize_document_value(row_model.doc_soma)
            found_norm = normalize_document_value(found_doc)
            auditoria = _compare_documents(row_model.doc_soma, found_doc)

            logger.warning("DOC_ENCONTRADO=%s", found_doc if found_doc is not None else "NENHUM")
            logger.warning("AUDITORIA=%s", auditoria)

            mark_row_auditoria(table, row_idx, auditoria)
            elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
            if auditoria == "Confirmado":
                log_kv(
                    logger,
                    "Conciliação concluída com sucesso.",
                    level=logging.INFO,
                    row=row_idx,
                    doc_esperado=expected_norm or "-",
                    doc_encontrado=found_norm or "NENHUM",
                    auditoria=auditoria,
                    elapsed_ms=elapsed_ms,
                )
                return AuditOutcome(analyzed=True, confirmed=True)

            log_kv(
                logger,
                "Conciliação divergente.",
                level=logging.WARNING,
                row=row_idx,
                doc_esperado=expected_norm or "-",
                doc_encontrado=found_norm or "NENHUM",
                auditoria=auditoria,
                elapsed_ms=elapsed_ms,
            )
            return AuditOutcome(analyzed=True, divergent=True)

    except Exception as e:
        elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
        logger.exception(
            "Erro tecnico na conciliacao | row=%s | err=%s | elapsed_ms=%s",
            row_idx,
            _safe_err(e),
            elapsed_ms,
        )
        return AuditOutcome(analyzed=True, technical_error=True)


def _build_summary(totals: AuditTotals) -> str:
    return (
        "\n" + "=" * 80 + "\n"
        "RESUMO DA CONCILIAÇÃO\n"
        + "=" * 80 + "\n"
        f"Linhas analisadas: {totals.analyzed}\n"
        f"Confirmados: {totals.confirmed}\n"
        f"Divergentes: {totals.divergent}\n"
        f"Erros técnicos: {totals.technical_errors}\n"
        + "=" * 80 + "\n"
    )


def _bootstrap_browser(settings: Any):
    bundle = WebDriverFactory.create(settings, headless=True)
    login = LoginPage(bundle.a, settings)
    if hasattr(login, "run"):
        login.run()
    else:
        login.login()
    return bundle


def main() -> int:
    overall_t0 = time.perf_counter()

    settings = _load_settings()
    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    ws = _sheet_name(settings)
    run_id = new_run_id()
    os.environ["RUN_ID"] = run_id

    _header()

    sheets = SheetsClient(settings)
    table = SheetsTable(sheets, ws)
    table.load()
    rows = _select_rows_to_audit(table)

    totals = AuditTotals()
    bundle = None

    if not rows:
        logger.warning("Nenhuma linha com AUDITORIA vazia para processar.")
    else:
        with step(logger, "audit.login_ui", run_id=run_id, sheet=ws):
            bundle = _bootstrap_browser(settings)

        try:
            page = EntradasSaidasPage(bundle.a, settings)

            for raw_row in rows:
                outcome = _audit_row(table=table, page=page, raw_row=raw_row, run_id=run_id)
                if not outcome.analyzed:
                    continue
                totals.analyzed += 1
                if outcome.confirmed:
                    totals.confirmed += 1
                elif outcome.divergent:
                    totals.divergent += 1
                elif outcome.technical_error:
                    totals.technical_errors += 1

        finally:
            try:
                if bundle is not None:
                    bundle.quit()
            except Exception:
                pass

    resumo_txt = _build_summary(totals)
    logger.warning(resumo_txt.replace("\n", " | "))
    print(resumo_txt)

    total_elapsed_ms = int((time.perf_counter() - overall_t0) * 1000)
    log_kv(
        logger,
        "Conciliacao finalizada.",
        level=logging.INFO,
        run_id=run_id,
        sheet=ws,
        analyzed=totals.analyzed,
        confirmed=totals.confirmed,
        divergent=totals.divergent,
        technical_errors=totals.technical_errors,
        elapsed_ms=total_elapsed_ms,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
