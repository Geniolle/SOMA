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
from soma_app.workflows.contaordem_writer import mark_row_doc_soma
from soma_app.workflows.process_contaordem import SheetsTable
from soma_app.workflows.run_soma import _load_settings, _sheet_name

logger = logging.getLogger(__name__)


def _safe_err(e: Exception) -> str:
    s = str(e).strip().replace("\n", " ")
    return s[:180] if s else type(e).__name__


def _is_blank(value: Any) -> bool:
    return not str(value or "").strip()


def _is_entrada_ou_saida(raw: Dict[str, Any]) -> bool:
    try:
        tipo = TipoMovimento.from_sheet_value(raw.get("TIPO", ""))
    except ValueError:
        return False
    return tipo in {TipoMovimento.ENTRADA, TipoMovimento.SAIDA}


@dataclass
class DocSomaTotals:
    eligible: int = 0
    searched: int = 0
    already_correct: int = 0
    corrected: int = 0
    not_found: int = 0
    technical_errors: int = 0


@dataclass
class DocSomaOutcome:
    searched: bool
    already_correct: bool = False
    corrected: bool = False
    not_found: bool = False
    technical_error: bool = False


def _header() -> None:
    logger.warning("=" * 80)
    logger.warning("CONCILIACAO DOC SOMA")
    logger.warning("=" * 80)


def _log_row_context(row: ContaOrdemRow) -> None:
    logger.warning("")
    logger.warning("LINHA=%s", row.row_number)
    logger.warning("ID_INTERNO=%s", row.id_interno)
    logger.warning("DATA=%s", row.data_mov)
    logger.warning("DESCRICAO=%s", row.descricao_soma)
    logger.warning("DOC_ATUAL=%s", row.doc_soma)


def _select_rows_to_conciliate(table: SheetsTable) -> list[Dict[str, Any]]:
    if not table.has_col("AUDITORIA"):
        raise RuntimeError("Coluna AUDITORIA nao encontrada na worksheet CONTAORDEM.")

    rows: list[Dict[str, Any]] = []
    skipped_auditoria = 0
    skipped_tipo = 0
    for raw in table.get_records_with_row():
        if not _is_blank(raw.get("AUDITORIA")):
            skipped_auditoria += 1
            continue
        if not _is_entrada_ou_saida(raw):
            skipped_tipo += 1
            continue
        rows.append(raw)

    if skipped_auditoria:
        logger.info("Linhas ignoradas por AUDITORIA preenchida: %s", skipped_auditoria)
    if skipped_tipo:
        logger.info("Linhas ignoradas por TIPO nao auditavel: %s", skipped_tipo)

    return rows


def _docs_equal(doc_atual: Any, doc_encontrado: Any) -> bool:
    return normalize_document_value(doc_atual) == normalize_document_value(doc_encontrado)


def _conciliate_row(
    *,
    table: SheetsTable,
    page: EntradasSaidasPage,
    raw_row: Dict[str, Any],
    run_id: str,
) -> DocSomaOutcome:
    row_idx = int(raw_row["row"])
    row_t0 = time.perf_counter()

    try:
        with step(logger, "doc_soma.process_row", run_id=run_id, row=row_idx):
            row_model = ContaOrdemRow.from_table_row(row_number=row_idx, raw=raw_row)
            _log_row_context(row_model)
            logger.warning("Pesquisando no SOMA...")

            found_doc = page.search_existing_doc(row_model)
            current_norm = normalize_document_value(row_model.doc_soma)
            found_norm = normalize_document_value(found_doc)

            if found_doc is None:
                logger.warning("DOC_ENCONTRADO=NENHUM")
                logger.warning("RESULTADO=NENHUM")
                logger.warning("ACAO=SEM_ALTERACAO")
                return DocSomaOutcome(searched=True, not_found=True)

            if _docs_equal(row_model.doc_soma, found_doc):
                logger.warning("DOC_ENCONTRADO=%s", found_doc)
                logger.warning("RESULTADO=IGUAL")
                logger.warning("ACAO=SEM_ALTERACAO")
                return DocSomaOutcome(searched=True, already_correct=True)

            logger.warning("DOC_ENCONTRADO=%s", found_doc)
            logger.warning("RESULTADO=DIFERENTE")
            mark_row_doc_soma(table, row_idx, found_norm)
            logger.warning("ACAO=DOC_SOMA_ATUALIZADO")
            logger.warning("NOVO_DOC_SOMA=%s", found_norm)
            logger.warning("DOC_ATUAL=%s", current_norm)
            return DocSomaOutcome(searched=True, corrected=True)

    except Exception as e:
        elapsed_ms = int((time.perf_counter() - row_t0) * 1000)
        logger.exception(
            "Erro tecnico na conciliacao DOC SOMA | row=%s | err=%s | elapsed_ms=%s",
            row_idx,
            _safe_err(e),
            elapsed_ms,
        )
        return DocSomaOutcome(searched=True, technical_error=True)


def _build_summary(totals: DocSomaTotals) -> str:
    return (
        "\n" + "=" * 80 + "\n"
        "RESUMO DA CONCILIACAO DOC SOMA\n"
        + "=" * 80 + "\n"
        f"Linhas elegiveis: {totals.eligible}\n"
        f"Pesquisadas: {totals.searched}\n"
        f"Documentos ja corretos: {totals.already_correct}\n"
        f"DOC. SOMA corrigidos: {totals.corrected}\n"
        f"Documentos nao encontrados: {totals.not_found}\n"
        f"Erros tecnicos: {totals.technical_errors}\n"
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
    rows = _select_rows_to_conciliate(table)

    totals = DocSomaTotals(eligible=len(rows))
    bundle = None

    if not rows:
        logger.warning("Nenhuma linha com AUDITORIA vazia para processar.")
    else:
        with step(logger, "doc_soma.login_ui", run_id=run_id, sheet=ws):
            bundle = _bootstrap_browser(settings)

        try:
            page = EntradasSaidasPage(bundle.a, settings)

            for raw_row in rows:
                totals.searched += 1
                outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id=run_id)
                if outcome.already_correct:
                    totals.already_correct += 1
                elif outcome.corrected:
                    totals.corrected += 1
                elif outcome.not_found:
                    totals.not_found += 1
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
        "Conciliacao DOC SOMA finalizada.",
        level=logging.INFO,
        run_id=run_id,
        sheet=ws,
        eligible=totals.eligible,
        searched=totals.searched,
        already_correct=totals.already_correct,
        corrected=totals.corrected,
        not_found=totals.not_found,
        technical_errors=totals.technical_errors,
        elapsed_ms=total_elapsed_ms,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
