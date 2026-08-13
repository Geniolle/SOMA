from __future__ import annotations

from soma_app.domain.reconciliation import ReconciliationCategory, reconcile_t_extrato


def _source(
    ident: str,
    *,
    doc: str = "5161859",
    data: str = "28/04/2026",
    tipo: str = "Entrada",
    descricao: str = "TRF.CRED  ALISSON RIOS",
    valor: str = "230,00",
    row: int = 10,
):
    return {
        "__row__": row,
        "DOC. SOMA": doc,
        "DATA MOV.": data,
        "DESCRIÇÃO": descricao,
        "IMPORTÂNCIA": valor,
        "TIPO": tipo,
        "STATUS": "Transferido",
        "ID_INTERNO": ident,
    }


def _conta(
    ident: str,
    *,
    doc: str = "5161859",
    data: str = "28/04/2026",
    tipo: str = "Entrada",
    descricao: str = "TRF.CREDALISSONRIOS",
    valor: str = "230,00",
    descricao_soma: str = "DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA) N002",
    processo: str = "T_EXTRATO",
    row: int = 20,
):
    return {
        "__row__": row,
        "DATA MOV.": data,
        "DESCRIÇÃO": descricao,
        "IMPORTÂNCIA": valor,
        "DOC. SOMA": doc,
        "TIPO": tipo,
        "DESCRIÇÃO SOMA": descricao_soma,
        "PROCESSO": processo,
        "ID_INTERNO": ident,
    }


def _soma(
    *,
    codigo: str = "5161859",
    tipo: str = "ENTRADA",
    descricao: str = "DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA) N002",
    valor: str = "230",
    pagamento: str = "28/04/2026",
    row: int = 30,
):
    return {
        "__row__": row,
        "CODIGO": codigo,
        "TIPO": tipo,
        "DESCRIÇÃO": descricao,
        "VALOR": valor,
        "PAGAMENTO": pagamento,
        "STATUS": "PAGO",
        "BAIXA": "SIM",
    }


def test_happy_path_normalizes_spacing_case_and_amount_sign():
    source = [_source("EXT0000002970", valor="-230,00")]
    conta = [_conta("EXT0000002970", valor="230")]
    soma = [_soma(valor="230,00")]

    report = reconcile_t_extrato(source, conta, soma, year=2026, month=4)

    assert report.source_rows_total == 1
    assert report.items[0].category == ReconciliationCategory.OK
    assert report.items[0].flow.endswith("SOMA ✅")


def test_missing_contaordem_is_separated():
    report = reconcile_t_extrato([_source("EXT0000003012")], [], [], year=2026, month=4)

    item = report.items[0]
    assert item.category == ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM
    assert any(issue.code == "CONTAORDEM_NAO_ENCONTRADA" for issue in item.issues)


def test_missing_soma_is_separated():
    report = reconcile_t_extrato(
        [_source("EXT0000002979", doc="5169999")],
        [_conta("EXT0000002979", doc="5169999")],
        [],
        year=2026,
        month=4,
    )

    item = report.items[0]
    assert item.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA
    assert any(issue.code == "SOMA_NAO_ENCONTRADO" for issue in item.issues)


def test_false_positive_like_ext2976_becomes_divergence():
    report = reconcile_t_extrato(
        [
            _source(
                "EXT0000002976",
                doc="5161859",
                tipo="Saída",
                descricao="PAG. CARTAO BUSTRADE",
                valor="-289,21",
            )
        ],
        [
            _conta(
                "EXT0000002976",
                doc="5161859",
                tipo="CARTAO",
                descricao="PAG.CARTAOBUSTRADE",
                valor="289,21",
                descricao_soma="PAGAMENTO CARTÃO BUSTRADE",
            )
        ],
        [_soma(codigo="5161859", tipo="ENTRADA", valor="230,00")],
        year=2026,
        month=4,
    )

    item = report.items[0]
    assert item.category == ReconciliationCategory.DIVERGENCIA
    codes = {issue.code for issue in item.issues}
    assert "TIPO_CONTA_SOMA" in codes
    assert "VALOR_CONTA_SOMA" in codes


def test_doc_reused_by_two_contaordem_ids_is_duplicate_for_affected_source():
    source = [_source("EXT0000002970")]
    conta = [
        _conta("EXT0000002970", doc="5161859", row=20),
        _conta("EXT0000002976", doc="5161859", row=21),
    ]
    soma = [_soma()]

    report = reconcile_t_extrato(source, conta, soma, year=2026, month=4)

    item = report.items[0]
    assert item.category == ReconciliationCategory.DUPLICADO
    assert any(issue.code == "DUP_DOC_CONTAORDEM" for issue in item.issues)


def test_excluded_cash_transfer_does_not_require_soma():
    source = [
        _source(
            "EXT0000002965",
            doc="Transferido",
            tipo="TRANSFERÊNCIA",
            descricao="ENT.NUMERARIO  CH24 0006774253",
            valor="3,00",
        )
    ]
    conta = [
        _conta(
            "EXT0000002965",
            doc="Transferido",
            tipo="TRANSFERÊNCIA",
            descricao="ENT.NUMERARIOCH240006774253",
            valor="3,00",
            descricao_soma="",
        )
    ]

    report = reconcile_t_extrato(source, conta, [], year=2026, month=4)

    item = report.items[0]
    assert item.category == ReconciliationCategory.EXCLUIDO
    assert item.excluded_from_soma is True
    assert item.flow.endswith("SOMA ⏭")


def test_source_doc_placeholder_is_warning_not_failure_when_conta_and_soma_match():
    report = reconcile_t_extrato(
        [_source("EXT0000003419", doc="Em processamento")],
        [_conta("EXT0000003419", doc="5161859")],
        [_soma(codigo="5161859")],
        year=2026,
        month=4,
    )

    item = report.items[0]
    assert item.category == ReconciliationCategory.OK
    assert any(warning.code == "ORIGEM_DOC_PENDENTE" for warning in item.warnings)


def test_period_filters_source_only():
    source = [
        _source("EXTAPR", data="28/04/2026"),
        _source("EXTMAY", data="01/05/2026", row=11),
    ]
    conta = [_conta("EXTAPR")]
    soma = [_soma()]

    report = reconcile_t_extrato(source, conta, soma, year=2026, month=4)

    assert [item.id_interno for item in report.items] == ["EXTAPR"]
    assert report.source_rows_total == 1
    assert report.source_rows_outside_period == 1


def test_real_source_doc_vs_placeholder_conta_is_reported_and_does_not_false_match_soma():
    report = reconcile_t_extrato(
        [
            _source(
                "EXT0000002976",
                doc="5161859",
                tipo="Cartão",
                descricao="PAG. CARTAO BUSTRADE",
                valor="-289,21",
            )
        ],
        [
            _conta(
                "EXT0000002976",
                doc="PGTO CARTÃO",
                tipo="Cartão",
                descricao="PAG.CARTAOBUSTRADE",
                valor="289,21",
                descricao_soma="PAGTO DE CARTÂO DE CRÉDITO",
            )
        ],
        [_soma(codigo="5161859", tipo="ENTRADA", valor="230,00")],
        year=2026,
        month=4,
    )

    item = report.items[0]
    assert item.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA
    codes = {issue.code for issue in item.issues}
    assert "DOC_ORIGEM_CONTA" in codes
    assert "CONTAORDEM_SEM_DOC_SOMA" in codes
    assert item.soma_rows == []
