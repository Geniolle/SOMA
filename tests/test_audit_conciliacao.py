from __future__ import annotations

from types import SimpleNamespace

from soma_app.domain.models import normalize_document_value
from soma_app.workflows.audit_conciliacao import (
    _audit_row,
    _bootstrap_browser,
    _compare_documents,
    _select_rows_to_audit,
)
from soma_app.workflows.contaordem_writer import mark_row_auditoria
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


def test_compare_documents_normalizes_numeric_text_and_float_repr():
    assert normalize_document_value(5165029) == "5165029"
    assert normalize_document_value("5165029") == "5165029"
    assert normalize_document_value("5165029.0") == "5165029"
    assert _compare_documents(5165029, "5165029.0") == "Confirmado"


def test_audit_row_confirms_matching_doc_and_writes_auditoria():
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

    outcome = _audit_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.analyzed is True
    assert outcome.confirmed is True
    assert outcome.divergent is False
    written = _written(fake, 2)
    assert written[table.col_idx("AUDITORIA")] == "Confirmado"


def test_audit_row_marks_divergent_when_found_doc_differs():
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
    page = FakePage("5165030")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165029", "AUDITORIA": ""}

    outcome = _audit_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.analyzed is True
    assert outcome.divergent is True
    written = _written(fake, 2)
    assert written[table.col_idx("AUDITORIA")] == "Divergente"


def test_audit_row_marks_divergent_when_no_doc_is_found():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0003",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage(None)
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165029", "AUDITORIA": ""}

    outcome = _audit_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.analyzed is True
    assert outcome.divergent is True
    written = _written(fake, 2)
    assert written[table.col_idx("AUDITORIA")] == "Divergente"


def test_audit_row_keeps_auditoria_empty_on_technical_error():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0004",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )
    page = FakePage(RuntimeError("timeout"))
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "5165029", "AUDITORIA": ""}

    outcome = _audit_row(table=table, page=page, raw_row=raw_row, run_id="t1")

    assert outcome.analyzed is True
    assert outcome.technical_error is True
    assert fake.updated_cells == []


def test_mark_row_auditoria_updates_only_auditoria_column():
    table, fake = _build_table(
        [
            {
                "TIPO": "Entrada",
                "DOC. SOMA": "5165029",
                "AUDITORIA": "",
                "ID_INTERNO": "EXT0005",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            }
        ]
    )

    mark_row_auditoria(table, 2, "Confirmado")

    written = _written(fake, 2)
    assert written == {table.col_idx("AUDITORIA"): "Confirmado"}


def test_select_rows_to_audit_skips_rows_with_auditoria_filled():
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
                "TIPO": "Entrada",
                "DOC. SOMA": "5165031",
                "AUDITORIA": "Divergente",
                "ID_INTERNO": "EXT0008",
                "DESCRIÇÃO SOMA": "Oferta culto",
                "DATA MOV.": "10/04/2026",
            },
        ]
    )

    rows = _select_rows_to_audit(table)

    assert [r["row"] for r in rows] == [2]


def test_bootstrap_browser_forces_headless_mode(monkeypatch):
    class FakeBundle:
        def __init__(self):
            self.a = object()
            self.quit_called = False

        def quit(self):
            self.quit_called = True

    captured = {}

    def fake_create(settings, *, headless=None, downloads_dir=None):
        captured["headless"] = headless
        return FakeBundle()

    class FakeLogin:
        def __init__(self, a, settings):
            self.a = a
            self.settings = settings
            captured["login"] = True

        def login(self):
            captured["login_method"] = "login"

    monkeypatch.setattr("soma_app.workflows.audit_conciliacao.WebDriverFactory.create", fake_create)
    monkeypatch.setattr("soma_app.workflows.audit_conciliacao.LoginPage", FakeLogin)

    bundle = _bootstrap_browser(SimpleNamespace())

    assert captured["headless"] is True
    assert captured["login"] is True
    assert captured["login_method"] == "login"
    assert isinstance(bundle, FakeBundle)
