from __future__ import annotations

from soma_app.domain.models import normalize_document_value
from soma_app.workflows.conciliacao_doc_soma import (
    _conciliate_row,
    _docs_equal,
    _select_rows_to_conciliate,
)
from soma_app.workflows.contaordem_writer import mark_row_doc_soma
from soma_app.workflows.process_contaordem import SheetsTable


class FakeSheetsClient:
    def __init__(self, header, records):
        self.header = header
        self.records = records
        self.updated_cells: list[tuple] = []
        self.batch_updates: list[tuple[str, list[tuple[str, list[list[object]]]]]] = []

    def get_header(self, ws: str, row: int = 1):
        return list(self.header)

    def get_all_records(self, ws: str):
        return [dict(r) for r in self.records]

    def batch_update(self, ws: str, ranges):
        payload = list(ranges)
        self.batch_updates.append((ws, payload))
        for a1, values in payload:
            cell = a1.split("!", 1)[-1]
            col = 0
            row_txt = ""
            for ch in cell:
                if ch.isalpha():
                    col = col * 26 + (ord(ch.upper()) - 64)
                elif ch.isdigit():
                    row_txt += ch
            row = int(row_txt) if row_txt else 0
            value = values[0][0] if values and values[0] else None
            self.updated_cells.append((ws, row, col, value))


class FakePage:
    def __init__(self, found_doc):
        self.found_doc = found_doc
        self.calls = []

    def search_existing_doc(self, row):
        self.calls.append(row)
        if isinstance(self.found_doc, Exception):
            raise self.found_doc
        return self.found_doc


def _build_table(records):
    header = ["TIPO", "DOC. SOMA", "AUDITORIA", "ID_INTERNO", "DESCRIÇÃO SOMA", "DATA MOV."]
    fake = FakeSheetsClient(header, records)
    table = SheetsTable(fake, "CONTAORDEM")
    table.load()
    return table, fake


def _written(fake, row):
    return {idx: value for (_ws, r, idx, value) in fake.updated_cells if r == row}


def test_docs_equal_handles_numeric_text_and_float_repr():
    assert normalize_document_value(5165029) == "5165029"
    assert normalize_document_value("5165029") == "5165029"
    assert normalize_document_value("5165029.0") == "5165029"
    assert _docs_equal(5165029, "5165029.0") is True


def test_conciliate_row_keeps_doc_soma_unchanged_when_equal():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0001",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage("5165029")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165029", "AUDITORIA": ""}

    outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.already_correct is True
    assert fake.updated_cells == []


def test_conciliate_row_updates_doc_soma_when_different():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0002",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage("5165044")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165029", "AUDITORIA": ""}

    outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.corrected is True
    written = _written(fake, 2)
    assert written == {table.col_idx("DOC. SOMA"): "5165044"}


def test_conciliate_row_updates_doc_soma_when_current_is_blank():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0003",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage("5165050")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "AUDITORIA": ""}

    outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.corrected is True
    written = _written(fake, 2)
    assert written == {table.col_idx("DOC. SOMA"): "5165050"}


def test_conciliate_row_keeps_doc_soma_when_not_found():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165051",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0004",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage(None)
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165051", "AUDITORIA": ""}

    outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.not_found is True
    assert fake.updated_cells == []


def test_conciliate_row_keeps_doc_soma_when_search_fails():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165052",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0005",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage(RuntimeError("timeout"))
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165052", "AUDITORIA": ""}

    outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.technical_error is True
    assert fake.updated_cells == []


def test_select_rows_to_conciliate_skips_rows_with_auditoria_filled():
    table, _fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0006",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            },
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165030",
                "AUDITORIA": "Confirmado",
                "ID_INTERNO": "EXT0007",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            },
            {
                "TIPO": "Transferência",
                "DOC. SOMA": "5165031",
                "AUDITORIA": "Divergente",
                "ID_INTERNO": "EXT0008",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            },
        ]
    )

    rows = _select_rows_to_conciliate(table)

    assert [r["row"] for r in rows] == [2]


def test_select_rows_to_conciliate_skips_tipo_cartao():
    table, _fake = _build_table(
        [
            {
                "TIPO": "Cartão",
                "DOC. SOMA": "5165000",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0011",
                "DESCRIÃ‡ÃƒO SOMA": "Compra cartão",
                "DATA MOV.": "10/04/2026",
            },
            {
                "TIPO": "Saída",
                "DOC. SOMA": "5165001",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0012",
                "DESCRIÃ‡ÃƒO SOMA": "Despesa",
                "DATA MOV.": "10/04/2026",
            },
        ]
    )

    rows = _select_rows_to_conciliate(table)

    assert [r["row"] for r in rows] == [3]


def test_mark_row_doc_soma_updates_only_doc_col():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0009",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )

    mark_row_doc_soma(table, 2, "DOC-9")

    written = _written(fake, 2)
    assert written == {table.col_idx("DOC. SOMA"): "DOC-9"}


def test_conciliate_row_accepts_simple_namespace_row_data():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0010",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage("5165029.0")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165029.0", "AUDITORIA": ""}

    outcome = _conciliate_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.already_correct is True
    assert fake.updated_cells == []
