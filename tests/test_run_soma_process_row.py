from __future__ import annotations

import logging
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


_UNSET = object()


class FakeEntradasSaidas:
    def __init__(
        self,
        *,
        doc_id: str = "DOC-E1",
        dados_doc: str = "",
        duplicate_doc: str | None = None,
        recover_doc: Any = _UNSET,
        fetch_exc: Exception | None = None,
        go_back_exc: Exception | None = None,
    ):
        self.doc_id = doc_id
        self.dados_doc = dados_doc
        self.duplicate_doc = duplicate_doc
        self.recover_doc = doc_id if recover_doc is _UNSET else recover_doc
        self.fetch_exc = fetch_exc
        self.go_back_exc = go_back_exc
        self.create_calls: list[Any] = []
        self.recover_calls: list[Any] = []
        self.precheck_calls: list[Any] = []
        self.fetch_calls: list[Any] = []
        self.go_back_calls: list[Any] = []

    def precheck_duplicate(self, row):
        self.precheck_calls.append(row)
        return self.duplicate_doc

    def create_and_get_doc_id(self, row):
        self.create_calls.append(row)
        return self.doc_id

    def recover_doc_id(self, row):
        self.recover_calls.append(row)
        return self.recover_doc

    def fetch_dados_doc(self, doc_id):
        self.fetch_calls.append(doc_id)
        if self.fetch_exc is not None:
            raise self.fetch_exc
        return self.dados_doc

    def go_back_to_list(self, row):
        self.go_back_calls.append(row)
        if self.go_back_exc is not None:
            raise self.go_back_exc


class FailingEntradasSaidas:
    def __init__(self, exc: Exception):
        self.exc = exc

    def precheck_duplicate(self, row):
        raise self.exc

    def create_and_get_doc_id(self, row):
        raise self.exc

    def recover_doc_id(self, row):
        raise self.exc

    def fetch_dados_doc(self, doc_id):
        return ""


def _search_doc_id(exc: Exception):
    # Nome da funcao importa: _is_pending_doc_exception detecta pelo nome do frame.
    raise exc


class PendingDocEntradasSaidas:
    def precheck_duplicate(self, row):
        return None

    def create_and_get_doc_id(self, row):
        _search_doc_id(RuntimeError("doc nao encontrado"))

    def recover_doc_id(self, row):
        _search_doc_id(RuntimeError("doc nao encontrado"))

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
    raw_row = {"row": 2, "TIPO": "Transferencia", "DOC. SOMA": "", "STATUS": ""}

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
    assert len(entradas_saidas.precheck_calls) == 1

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DOC-E1"


def test_process_row_marks_ok_before_optional_post_create_sync(monkeypatch):
    table, fake = _build_table()
    entradas_saidas = FakeEntradasSaidas(
        doc_id="DOC-E2",
        fetch_exc=RuntimeError("falha ao ler dados doc"),
        go_back_exc=RuntimeError("falha ao voltar"),
    )
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    monkeypatch.setenv("SOMA_SYNC_DOC_AFTER_CREATE", "1")

    outcome = _call(table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias())

    assert outcome.ok is True
    assert outcome.created is True
    assert len(entradas_saidas.create_calls) == 1
    assert len(entradas_saidas.fetch_calls) == 1
    assert len(entradas_saidas.go_back_calls) == 1

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DOC-E2"
    assert written[table.col_idx("STATUS")] == "VALIDADO"


def test_process_row_logs_step_timings_for_success_path(monkeypatch, caplog):
    table, fake = _build_table()
    entradas_saidas = FakeEntradasSaidas(doc_id="DOC-E3")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    monkeypatch.delenv("SOMA_SYNC_DOC_AFTER_CREATE", raising=False)

    with caplog.at_level(logging.INFO):
        outcome = _call(table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias())

    assert outcome.ok is True
    assert "Tempo do subpasso da linha" in caplog.text
    assert "stage=precheck_duplicate" in caplog.text
    assert "stage=create_and_get_doc_id" in caplog.text
    assert "stage=post_create_sync" in caplog.text
    assert "stage=row_total" in caplog.text


def test_process_row_marks_duplicate_without_creating_doc():
    table, fake = _build_table()
    entradas_saidas = FakeEntradasSaidas(duplicate_doc="123456")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias())

    assert outcome.ok is True
    assert outcome.duplicated is True
    assert outcome.created is False
    assert outcome.recovered is False
    assert len(entradas_saidas.precheck_calls) == 1
    assert len(entradas_saidas.create_calls) == 0

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DUPLICADO"
    assert written[table.col_idx("STATUS")] == "DUPLICADO"


def test_process_row_pending_doc_still_creates_when_not_duplicate():
    table, fake = _build_table()
    entradas_saidas = FakeEntradasSaidas(doc_id="DOC-R1")
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "EM ERRO", "STATUS": "PENDENTE_DOC"}

    outcome = _call(table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias())

    assert outcome.ok is True
    assert outcome.created is True
    assert outcome.recovered is False
    assert len(entradas_saidas.recover_calls) == 0
    assert len(entradas_saidas.create_calls) == 1
    assert len(entradas_saidas.precheck_calls) == 1

    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == "DOC-R1"
    assert written[table.col_idx("STATUS")] == "VALIDADO"


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
    entradas_saidas = FailingEntradasSaidas(RuntimeError("falha generica"))
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
    entradas_saidas = FailingEntradasSaidas(RuntimeError("falha generica"))
    raw_row = {"row": 2, "TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}

    outcome = _call(
        table, raw_row, entradas_saidas=entradas_saidas, transferencias=FakeTransferencias(), allow_retry=True
    )

    assert outcome.ok is False
    written = _written(fake, 2)
    assert written[table.col_idx("DOC. SOMA")] == ""
    assert written[table.col_idx("STATUS")] == "VALIDADO"
