from __future__ import annotations

from typing import Any

from soma_app.workflows.process_contaordem import SheetsTable
from soma_app.workflows.run_soma import _process_row


class FakeSheetsClient:
    def __init__(self, header, records):
        self.header = header
        self.records = records
        self.updated_cells: list[tuple] = []

    def get_header(self, ws: str, row: int = 1):
        return list(self.header)

    def get_all_records(self, ws: str):
        return [dict(r) for r in self.records]

    def update_cell(self, ws: str, row: int, col: int, value):
        self.updated_cells.append((ws, row, col, value))


class FakeTransferencias:
    def __init__(self, doc_id: str = "DOC-T1"):
        self.doc_id = doc_id
        self.calls: list[Any] = []

    def run(self, row):
        self.calls.append(row)
        return self.doc_id


class FakeEntradasSaidas:
    def __init__(self, *, doc_id: str = "DOC-E1", dados_doc: str = ""):
        self.doc_id = doc_id
        self.dados_doc = dados_doc
        self.create_calls: list[Any] = []
        self.recover_calls: list[Any] = []

    def create_and_get_doc_id(self, row):
        self.create_calls.append(row)
        return self.doc_id

    def recover_doc_id(self, row):
        self.recover_calls.append(row)
        return self.doc_id

    def fetch_dados_doc(self, doc_id):
        return self.dados_doc


class FailingEntradasSaidas:
    def __init__(self, exc: Exception):
        self.exc = exc

    def create_and_get_doc_id(self, row):
        raise self.exc

    def recover_doc_id(self, row):
        raise self.exc

    def fetch_dados_doc(self, doc_id):
        return ""


def _search_doc_id(exc: Exception):
    # Nome da função importa: _is_pending_doc_exception detecta pelo nome do frame.
    raise exc


class PendingDocEntradasSaidas:
    def create_and_get_doc_id(self, row):
        _search_doc_id(RuntimeError("doc não encontrado"))

    def recover_doc_id(self, row):
        _search_doc_id(RuntimeError("doc não encontrado"))

    def fetch_dados_doc(self, doc_id):
        return ""


def _build_table():
    header = ["TIPO", "DOC. SOMA", "STATUS", "IDUSER", "TIMESTAMP"]
    records = [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": "", "IDUSER": "", "TIMESTAMP": ""}]
    fake = FakeSheetsClient(header, records)
    table = SheetsTable(fake, "CONTAORDEM")
    table.load()
    return table, fake


def _written(fake, row):
    return {idx: value for (_ws, r, idx, value) in fake.updated_cells if r == row}


def _call(table, raw_row, *, entradas_saidas, transferencias, allow_retry=False):
    return _process_row(
        table,
        raw_row,
        run_id="t1",
        batch=1,
        progress_current=1,
        progress_total=1,
        iduser="USERJOB",
        allow_retry=allow_retry,
        entradas_saidas=entradas_saidas,
        transferencias=transferencias,
    )


def test_process_row_transferencia_marks_ok():
    table, fake = _build_table()
    transferencias = FakeTransferencias(doc_id="DOC-T1")
    raw_row = {"row": 2, "TIPO": "Transferência", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(table, raw_row, entradas_saidas=FakeEntradasSaidas(), transferencias=transferencias)

    assert outcome.ok is True
    assert outcome.transfer is True
    assert outcome.created is False
    assert outcome.recovered is False
    assert len(transferencias.calls) == 1

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DOC-T1"
    assert written[table.col_idx("STATUS")] == "VALIDADO"


def test_process_row_creates_doc_for_new_entrada():
    table, fake = _build_table()
    entradas_saidas = FakeEntradasSaidas(doc_id="DOC-E1")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias())

    assert outcome.ok is True
    assert outcome.created is True
    assert outcome.recovered is False
    assert len(entradas_saidas.create_calls) == 1
    assert len(entradas_saidas.recover_calls) == 0

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DOC-E1"


def test_process_row_recovers_doc_for_pending_doc_status():
    table, fake = _build_table()
    entradas_saidas = FakeEntradasSaidas(doc_id="DOC-R1")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "EM ERRO", "STATUS": "PENDENTE_DOC"}

    outcome = _call(table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias())

    assert outcome.ok is True
    assert outcome.recovered is True
    assert outcome.created is False
    assert len(entradas_saidas.recover_calls) == 1
    assert len(entradas_saidas.create_calls) == 0


def test_process_row_pending_doc_exception_forces_pendente_doc_status():
    table, fake = _build_table()
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(
        table, raw_row, entradas_saidas=PendingDocEntradasSaidas(), transferencias=FakeTransferencias()
    )

    assert outcome.ok is False
    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "EM ERRO"
    assert written[table.col_idx("STATUS")] == "PENDENTE_DOC"


def test_process_row_generic_exception_marks_em_erro_when_no_retry():
    table, fake = _build_table()
    entradas_saidas = FailingEntradasSaidas(RuntimeError("falha genérica"))
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(
        table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias(), allow_retry=False
    )

    assert outcome.ok is False
    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "EM ERRO"
    assert written[table.col_idx("STATUS")] == "ERRO"


def test_process_row_generic_exception_clears_doc_when_allow_retry():
    table, fake = _build_table()
    entradas_saidas = FailingEntradasSaidas(RuntimeError("falha genérica"))
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(
        table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias(), allow_retry=True
    )

    assert outcome.ok is False
    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == ""
    assert written[table.col_idx("STATUS")] == "VALIDADO"
