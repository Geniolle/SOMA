from __future__ import annotations

import pytest

from soma_app.workflows.process_contaordem import SheetsTable
from soma_app.workflows.process_soma import (
    _existing_codes,
    _filtrar_novos_lancamentos,
    _is_target_row_blank,
    _next_write_row,
    _norm_code,
    _norm_text,
    _resolve_col_name,
)


class FakeSheetsClient:
    def __init__(self, header, records):
        self.header = header
        self.records = records

    def get_header(self, ws: str, row: int = 1):
        return list(self.header)

    def get_all_records(self, ws: str):
        return [dict(r) for r in self.records]


def _build_table(header, records):
    table = SheetsTable(FakeSheetsClient(header, records), "SOMA")
    table.load()
    return table


def test_norm_text_strips_accents_and_case():
    assert _norm_text("  Descrição  ") == "descricao"


def test_norm_code_strips_trailing_float_zero():
    assert _norm_code("123.0") == "123"
    assert _norm_code("123.00") == "123"
    assert _norm_code("ABC-1") == "ABC-1"
    assert _norm_code(None) == ""


def test_is_target_row_blank():
    assert _is_target_row_blank({"A": "", "B": None}, ["A", "B"]) is True
    assert _is_target_row_blank({"A": "x", "B": None}, ["A", "B"]) is False


def test_next_write_row_uses_first_blank_row():
    table = _build_table(["CODIGO"], [{"CODIGO": "1"}, {"CODIGO": ""}, {"CODIGO": ""}])
    assert _next_write_row(table, "CODIGO", ["CODIGO"]) == 3


def test_next_write_row_falls_back_to_end_when_hole_has_data_below():
    table = _build_table(["CODIGO"], [{"CODIGO": "1"}, {"CODIGO": ""}, {"CODIGO": "2"}])
    assert _next_write_row(table, "CODIGO", ["CODIGO"]) == 5


def test_next_write_row_appends_when_no_blank_row():
    table = _build_table(["CODIGO"], [{"CODIGO": "1"}, {"CODIGO": "2"}])
    assert _next_write_row(table, "CODIGO", ["CODIGO"]) == 4


def test_existing_codes_normalizes():
    table = _build_table(["CODIGO"], [{"CODIGO": "1.0"}, {"CODIGO": "2"}, {"CODIGO": ""}])
    assert _existing_codes(table, "CODIGO") == {"1", "2"}


def test_resolve_col_name_matches_alias_ignoring_accents():
    table = _build_table(["Código", "Tipo"], [])
    assert _resolve_col_name(table, "CODIGO") == "Código"


def test_resolve_col_name_raises_when_missing():
    table = _build_table(["Tipo"], [])
    with pytest.raises(RuntimeError):
        _resolve_col_name(table, "CODIGO")


def test_filtrar_novos_lancamentos_dedupes_existing_and_batch_duplicates():
    lancamentos = [
        {"codigo": "1", "tipo": "Entrada"},
        {"codigo": "2", "tipo": "Saída"},
        {"codigo": "2", "tipo": "Saída"},
        {"codigo": "", "tipo": "Entrada"},
    ]
    novos, ignorados_existente, ignorados_lote = _filtrar_novos_lancamentos(lancamentos, existing_codes={"1"})
    assert [n["codigo"] for n in novos] == ["2"]
    assert ignorados_existente == 1
    assert ignorados_lote == 1
