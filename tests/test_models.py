from __future__ import annotations

from soma_app.domain.models import LinhaStatus, TipoMovimento, status_from_doc_soma


def test_tipo_movimento_from_sheet_value():
    assert TipoMovimento.from_sheet_value("entrada") is TipoMovimento.ENTRADA
    assert TipoMovimento.from_sheet_value("Saida") is TipoMovimento.SAIDA
    assert TipoMovimento.from_sheet_value("transferencia") is TipoMovimento.TRANSFERENCIA


def test_status_from_doc_soma():
    assert status_from_doc_soma("") is LinhaStatus.PENDENTE
    assert status_from_doc_soma("Em processamento") is LinhaStatus.EM_PROCESSAMENTO
    assert status_from_doc_soma("Em erro") is LinhaStatus.EM_ERRO
    assert status_from_doc_soma("12345") is LinhaStatus.CONCLUIDA
