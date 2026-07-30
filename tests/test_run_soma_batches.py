from __future__ import annotations

from soma_app.workflows import run_soma as run_soma_module
from soma_app.workflows.process_contaordem import preprocess_contaordem


class FakeSheetsClient:
    """
    Não persiste locks nos records (ao contrário do gspread real). Isso serve
    para provar que _run_batches para de reprocessar via attempted_rows,
    e não porque o backend refletiu o lock gravado.
    """

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


class FakeEntradasSaidas:
    def __init__(self, doc_id: str = "DOC-1"):
        self.doc_id = doc_id
        self.created: list = []

    def create_and_get_doc_id(self, row):
        self.created.append(row)
        return self.doc_id

    def recover_doc_id(self, row):
        raise AssertionError("recover_doc_id não deveria ser chamado neste teste")

    def fetch_dados_doc(self, doc_id):
        return ""


class FakeTransferencias:
    def run(self, row):
        raise AssertionError("transferencias.run não deveria ser chamado neste teste")


def test_run_batches_processes_workset_once_and_stops_via_attempted_rows(monkeypatch):
    header = ["TIPO", "DOC. SOMA", "STATUS", "IDUSER", "TIMESTAMP"]
    records = [{"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""}]
    fake_sheets = FakeSheetsClient(header, records)

    monkeypatch.setattr(run_soma_module, "SheetsClient", lambda settings: fake_sheets)

    result = preprocess_contaordem(fake_sheets, ws="CONTAORDEM", run_id="t1", batch=1)
    assert len(result.workset) == 1

    entradas_saidas = FakeEntradasSaidas(doc_id="DOC-1")
    transferencias = FakeTransferencias()

    totals = run_soma_module._run_batches(
        settings=None,
        ws="CONTAORDEM",
        run_id="t1",
        sheets=fake_sheets,
        result=result,
        batch=1,
        bundle=None,
        iduser="USERJOB",
        allow_retry=False,
        run_caixas_bancos=False,
        entradas_saidas=entradas_saidas,
        transferencias=transferencias,
    )

    assert totals.processed == 1
    assert totals.ok == 1
    assert totals.created == 1
    assert totals.err == 0
    assert len(totals.row_times_ms) == 1
    assert len(entradas_saidas.created) == 1


def test_run_batches_counts_errors_without_stopping(monkeypatch):
    header = ["TIPO", "DOC. SOMA", "STATUS"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
        {"TIPO": "Transferência", "DOC. SOMA": "", "STATUS": ""},
    ]
    fake_sheets = FakeSheetsClient(header, records)
    monkeypatch.setattr(run_soma_module, "SheetsClient", lambda settings: fake_sheets)

    result = preprocess_contaordem(fake_sheets, ws="CONTAORDEM", run_id="t1", batch=1)
    assert len(result.workset) == 1  # prioridade: Entrada primeiro, Transferência fica para o próximo batch

    class BoomEntradasSaidas(FakeEntradasSaidas):
        def create_and_get_doc_id(self, row):
            raise RuntimeError("falha ao criar doc")

    totals = run_soma_module._run_batches(
        settings=None,
        ws="CONTAORDEM",
        run_id="t1",
        sheets=fake_sheets,
        result=result,
        batch=1,
        bundle=None,
        iduser="USERJOB",
        allow_retry=False,
        run_caixas_bancos=False,
        entradas_saidas=BoomEntradasSaidas(),
        transferencias=FakeTransferencias(),
    )

    assert totals.processed == 1
    assert totals.err == 1
    assert totals.ok == 0


def test_run_batches_can_stop_after_max_rows_per_run(monkeypatch):
    header = ["TIPO", "DOC. SOMA", "STATUS", "IDUSER", "TIMESTAMP", "DADOS DOC"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
    ]
    fake_sheets = FakeSheetsClient(header, records)
    monkeypatch.setattr(run_soma_module, "SheetsClient", lambda settings: fake_sheets)

    result = preprocess_contaordem(fake_sheets, ws="CONTAORDEM", run_id="t1", batch=1)
    assert len(result.workset) == 2

    entradas_saidas = FakeEntradasSaidas(doc_id="DOC-1")
    transferencias = FakeTransferencias()

    totals = run_soma_module._run_batches(
        settings=None,
        ws="CONTAORDEM",
        run_id="t1",
        sheets=fake_sheets,
        result=result,
        batch=1,
        bundle=None,
        iduser="USERJOB",
        allow_retry=False,
        run_caixas_bancos=False,
        max_rows_per_run=1,
        entradas_saidas=entradas_saidas,
        transferencias=transferencias,
    )

    assert totals.processed == 1
    assert totals.ok == 1
    assert len(entradas_saidas.created) == 1
