from __future__ import annotations

from soma_app.workflows.process_contaordem import preprocess_contaordem


class FakeSheetsClient:
    def __init__(self, header, records):
        self.header = header
        self.records = records
        self.updated_cells = []

    def get_header(self, ws: str, row: int = 1):
        return list(self.header)

    def get_all_records(self, ws: str):
        return [dict(r) for r in self.records]

    def update_cell(self, ws: str, row: int, col: int, value):
        self.updated_cells.append((ws, row, col, value))


def test_locked_row_is_candidate_for_reprocessing():
    header = ["TIPO", "DOC. SOMA", "STATUS"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "Em processamento", "STATUS": ""},
    ]
    sheets = FakeSheetsClient(header=header, records=records)

    result = preprocess_contaordem(
        sheets,
        ws="CONTAORDEM",
        run_id="test",
        batch=1,
        batch_size=1,
    )

    assert len(result.workset) == 1
    assert int(result.workset[0]["row"]) == 2
    # A linha já está lockada; não deve gerar novo update de lock.
    assert result.newly_locked == 0


def test_inflight_does_not_reduce_batch_capacity_for_reprocess():
    header = ["TIPO", "DOC. SOMA", "STATUS"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "Em processamento", "STATUS": ""},
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
    ]
    sheets = FakeSheetsClient(header=header, records=records)

    result = preprocess_contaordem(
        sheets,
        ws="CONTAORDEM",
        run_id="test",
        batch=1,
        batch_size=1,
    )

    # Com batch=1, ainda deve processar 1 linha (não ficar 0 por causa de inflight).
    assert len(result.workset) == 1


def test_status_is_ignored_for_queue_validation():
    header = ["TIPO", "DOC. SOMA", "STATUS"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": "ERRO"},
    ]
    sheets = FakeSheetsClient(header=header, records=records)

    result = preprocess_contaordem(
        sheets,
        ws="CONTAORDEM",
        run_id="test",
        batch=1,
        batch_size=5,
    )

    assert len(result.workset) == 1
    assert int(result.workset[0]["row"]) == 2


def test_em_erro_doc_is_candidate_for_recovery():
    header = ["TIPO", "DOC. SOMA", "STATUS"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "EM ERRO", "STATUS": ""},
    ]
    sheets = FakeSheetsClient(header=header, records=records)

    result = preprocess_contaordem(
        sheets,
        ws="CONTAORDEM",
        run_id="test",
        batch=1,
        batch_size=5,
    )

    assert len(result.workset) == 1
    assert int(result.workset[0]["row"]) == 2


def test_eligible_total_counts_all_docs_available_before_batch_cut():
    header = ["TIPO", "DOC. SOMA", "STATUS"]
    records = [
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
        {"TIPO": "Entrada", "DOC. SOMA": "", "STATUS": ""},
    ]
    sheets = FakeSheetsClient(header=header, records=records)

    result = preprocess_contaordem(
        sheets,
        ws="CONTAORDEM",
        run_id="test",
        batch=1,
        batch_size=2,
    )

    assert result.eligible_total == 3
    assert len(result.workset) == 2
