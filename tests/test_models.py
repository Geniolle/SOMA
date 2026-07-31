from __future__ import annotations

from soma_app.domain.models import ContaOrdemRow, LinhaStatus, TipoMovimento, status_from_doc_soma


def test_tipo_movimento_from_sheet_value():
    assert TipoMovimento.from_sheet_value("entrada") is TipoMovimento.ENTRADA
    assert TipoMovimento.from_sheet_value("Saida") is TipoMovimento.SAIDA
    assert TipoMovimento.from_sheet_value("transferencia") is TipoMovimento.TRANSFERENCIA


def test_status_from_doc_soma():
    assert status_from_doc_soma("") is LinhaStatus.PENDENTE
    assert status_from_doc_soma("Em processamento") is LinhaStatus.EM_PROCESSAMENTO
    assert status_from_doc_soma("Em erro") is LinhaStatus.EM_ERRO
    assert status_from_doc_soma("12345") is LinhaStatus.CONCLUIDA


def test_conta_ordem_row_loads_id_interno():
    raw = {
        "TIPO": "Saída",
        "DATA MOV.": "2024-01-15",
        "CAIXA": "CAIXA PRINCIPAL",
        "CAIXA SAIDA": "CAIXA SAIDA",
        "CENTRO DE CUSTO": "CC001",
        "PLANO DE CONTA": "PLANO001",
        "FORMA DE PAGAMENTO": "Transferência Bancária",
        "IMPORTÂNCIA": "1000.00",
        "DESCRIÇÃO SOMA": "Descrição teste",
        "ID_INTERNO": "INT123456",
        "DOC. SOMA": "",
        "DADOS DOC": "DADOS_TESTE",
        "IDUSER": "USER1",
        "TIMESTAMP": "2024-01-15 10:00:00",
    }
    row = ContaOrdemRow.from_table_row(row_number=1, raw=raw)
    assert row.id_interno == "INT123456"


def test_conta_ordem_row_loads_id_interno_with_alternate_header():
    raw = {
        "TIPO": "Saída",
        "DATA MOV.": "2024-01-15",
        "CAIXA": "CAIXA PRINCIPAL",
        "CAIXA SAIDA": "CAIXA SAIDA",
        "CENTRO DE CUSTO": "CC001",
        "PLANO DE CONTA": "PLANO001",
        "FORMA DE PAGAMENTO": "Transferência Bancária",
        "IMPORTÂNCIA": "1000.00",
        "DESCRIÇÃO SOMA": "Descrição teste",
        "ID INTERNO": "INT654321",
        "DOC. SOMA": "",
        "DADOS DOC": "DADOS_TESTE",
        "IDUSER": "USER1",
        "TIMESTAMP": "2024-01-15 10:00:00",
    }
    row = ContaOrdemRow.from_table_row(row_number=2, raw=raw)
    assert row.id_interno == "INT654321"


def test_conta_ordem_row_id_interno_empty_when_missing():
    raw = {
        "TIPO": "Entrada",
        "DATA MOV.": "2024-01-15",
        "CAIXA": "CAIXA PRINCIPAL",
        "CAIXA SAIDA": "",
        "CENTRO DE CUSTO": "CC001",
        "PLANO DE CONTA": "PLANO001",
        "FORMA DE PAGAMENTO": "Dinheiro",
        "IMPORTÂNCIA": "500.00",
        "DESCRIÇÃO SOMA": "Descrição teste",
        "DOC. SOMA": "",
        "DADOS DOC": "DADOS_TESTE",
        "IDUSER": "USER1",
        "TIMESTAMP": "2024-01-15 10:00:00",
    }
    row = ContaOrdemRow.from_table_row(row_number=3, raw=raw)
    assert row.id_interno == ""
