from __future__ import annotations

import logging
from datetime import datetime
from typing import Any, Dict, Optional

from soma_app.workflows.process_contaordem import (
    DOC_COL_DEFAULT,
    LOCK_VALUE_DEFAULT,
    STATUS_COL_DEFAULT,
    SheetsTable,
)

logger = logging.getLogger(__name__)


def _now_pt() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def _time_col_name(t: SheetsTable) -> Optional[str]:
    for c in ("TEMPO_PROCESSAMENTO_MS", "TEMPO_LINHA_MS", "PROCESS_TIME_MS", "DURACAO_MS"):
        if t.has_col(c):
            return c
    return None


def mark_row_ok(
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
    logger.info("DOC. SOMA gravado na sheet | row=%s | doc=%s | status=VALIDADO", row_idx, doc_id)


def mark_row_duplicate(
    t: SheetsTable,
    row_idx: int,
    iduser: str,
    *,
    duplicate_doc: str | None = None,
    elapsed_ms: int | None = None,
) -> None:
    updates = [
        (row_idx, "IDUSER", iduser),
        (row_idx, "TIMESTAMP", _now_pt()),
    ]
    if duplicate_doc is not None:
        updates.insert(0, (row_idx, DOC_COL_DEFAULT, duplicate_doc))
    else:
        updates.insert(0, (row_idx, DOC_COL_DEFAULT, ""))
    if t.has_col(STATUS_COL_DEFAULT):
        updates.append((row_idx, STATUS_COL_DEFAULT, "DUPLICADO"))
    if elapsed_ms is not None:
        c = _time_col_name(t)
        if c:
            updates.append((row_idx, c, elapsed_ms))
    t.batch_update_cells(updates)


def mark_row_auditoria(
    t: SheetsTable,
    row_idx: int,
    auditoria: str,
) -> None:
    t.batch_update_cells([(row_idx, "AUDITORIA", auditoria)])


def mark_row_doc_soma(
    t: SheetsTable,
    row_idx: int,
    doc_soma: str,
) -> None:
    t.batch_update_cells([(row_idx, DOC_COL_DEFAULT, doc_soma)])


def mark_row_error(
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


def unlock_still_processing(t: SheetsTable, run_rows: Dict[int, Dict[str, Any]]) -> None:
    updates = []
    for row_idx, r in run_rows.items():
        doc = str(r.get(DOC_COL_DEFAULT, "")).strip()
        if doc.lower() == LOCK_VALUE_DEFAULT.lower():
            updates.append((row_idx, DOC_COL_DEFAULT, ""))
            if t.has_col(STATUS_COL_DEFAULT):
                updates.append((row_idx, STATUS_COL_DEFAULT, "VALIDADO"))
    if updates:
        t.batch_update_cells(updates)
