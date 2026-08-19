from __future__ import annotations

from soma_app.workflows.contaordem_writer import (
    mark_row_duplicate,
    mark_row_error,
    mark_row_ok,
    unlock_still_processing,
)
from soma_app.workflows.process_contaordem import SheetsTable


class FakeSheetsClient:
    def __init__(self, header, records):
        self.header = header
        self.records = records
        self.updated_cells: list[tuple] = []
        self.header_reads = 0
        self.batch_updates: list[tuple[str, list[tuple[str, list[list[object]]]]]] = []

    def get_header(self, ws: str, row: int = 1):
        self.header_reads += 1
        return list(self.header)

    def get_all_records(self, ws: str):
        return [dict(r) for r in self.records]

    def update_cell(self, ws: str, row: int, col: int, value):
        self.updated_cells.append((ws, row, col, value))

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


def _build_table(header, records):
    fake = FakeSheetsClient(header, records)
    table = SheetsTable(fake, "CONTAORDEM")
    table.load()
    return table, fake


def _written(fake, row):
    return {idx: value for (_ws, r, idx, value) in fake.updated_cells if r == row}


def test_mark_row_ok_sets_doc_status_iduser_and_timestamp():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS", "IDUSER", "TIMESTAMP"],
        [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": "", "IDUSER": "", "TIMESTAMP": ""}],
    )

    mark_row_ok(table, 2, "DOC-1", "USERJOB")

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DOC-1"
    assert written[table.col_idx("STATUS")] == "VALIDADO"
    assert written[table.col_idx("IDUSER")] == "USERJOB"
    assert table.col_idx("TIMESTAMP") in written


def test_mark_row_error_defaults_to_em_erro_when_not_allow_retry():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS"],
        [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}],
    )

    mark_row_error(table, 2, "falha", allow_retry=False)

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "EM ERRO"
    assert written[table.col_idx("STATUS")] == "ERRO"


def test_mark_row_error_clears_doc_when_allow_retry():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS"],
        [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}],
    )

    mark_row_error(table, 2, "falha", allow_retry=True)

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == ""
    assert written[table.col_idx("STATUS")] == "VALIDADO"


def test_mark_row_error_force_doc_and_status_override_allow_retry():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS"],
        [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}],
    )

    mark_row_error(table, 2, "falha", allow_retry=False, force_doc="EM ERRO", force_status="PENDENTE_DOC")

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "EM ERRO"
    assert written[table.col_idx("STATUS")] == "PENDENTE_DOC"


def test_mark_row_duplicate_sets_duplicate_doc_and_status():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS", "IDUSER", "TIMESTAMP"],
        [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": "", "IDUSER": "", "TIMESTAMP": ""}],
    )

    mark_row_duplicate(table, 2, "USERJOB", duplicate_doc="123456")

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "123456"
    assert written[table.col_idx("STATUS")] == "DUPLICADO"
    assert written[table.col_idx("IDUSER")] == "USERJOB"
    assert table.col_idx("TIMESTAMP") in written


def test_unlock_still_processing_resets_locked_rows_only():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS"],
        [
            {"TIPO": "Entrada", "DOC. SOMA": "Em processamento", "STATUS": ""},
            {"TIPO": "Entrada", "DOC. SOMA": "DOC-9", "STATUS": "VALIDADO"},
        ],
    )
    run_rows = {2: table.records[0], 3: table.records[1]}

    unlock_still_processing(table, run_rows)

    written_row2 = _written(fake, 2)
    assert written_row2[table.col_idx("DOC. SOMA")] == ""
    assert written_row2[table.col_idx("STATUS")] == "VALIDADO"
    assert all(r != 3 for (_ws, r, _c, _v) in fake.updated_cells)


def test_batch_update_cells_prefers_direct_range_updates_without_reloading_header():
    table, fake = _build_table(
        ["TIPO", "DOC. SOMA", "STATUS", "IDUSER", "TIMESTAMP"],
        [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": "", "IDUSER": "", "TIMESTAMP": ""}],
    )

    fake.updated_cells.clear()
    fake.batch_updates.clear()
    fake.header_reads = 0

    table.batch_update_cells([(2, "DOC. SOMA", "DOC-9"), (2, "STATUS", "VALIDADO")])

    assert fake.header_reads == 0
    assert fake.batch_updates == [
        (
            "CONTAORDEM",
            [
                ("CONTAORDEM!B2", [["DOC-9"]]),
                ("CONTAORDEM!C2", [["VALIDADO"]]),
            ],
        )
    ]
    assert fake.updated_cells == [
        ("CONTAORDEM", 2, table.col_idx("DOC. SOMA"), "DOC-9"),
        ("CONTAORDEM", 2, table.col_idx("STATUS"), "VALIDADO"),
    ]
