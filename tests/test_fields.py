from soma_app.config.fields import (
    ENTRADA_COMMON_FIELDS,
    ENTRADA_SPECIFIC_FIELDS,
    FORM_FIELD_REGISTRY,
    CaixaBancoField,
    EntradaSaidaField,
    TransferenciaField,
    field_names,
)


def test_entrada_field_names_are_semantic_and_stable() -> None:
    assert field_names(ENTRADA_COMMON_FIELDS) == ("PLANO_CONTA", "CENTRO_CUSTO", "DESCRICAO", "VALOR", "OBS")
    assert field_names(ENTRADA_SPECIFIC_FIELDS) == ("DATA_ENTRADA", "FORMA_PAGAMENTO_ENTRADA", "CAIXA_ENTRADA")
    assert EntradaSaidaField.PLANO_CONTA.value == "PLANO_CONTA"


def test_transferencia_and_caixas_fields_are_stable() -> None:
    assert TransferenciaField.CAIXA_SAIDA.value == "CAIXA_SAIDA"
    assert TransferenciaField.CAIXA_ENTRADA.value == "CAIXA_ENTRADA"
    assert CaixaBancoField.CAIXA_DIARIO.value == "CAIXA_DIARIO"
    assert CaixaBancoField.LIVRARIA.value == "VERBO_SHOP"


def test_form_field_registry_groups_are_consistent() -> None:
    assert FORM_FIELD_REGISTRY.entradas_saidas.common == ENTRADA_COMMON_FIELDS
    assert FORM_FIELD_REGISTRY.entradas_saidas.entrada_specific == ENTRADA_SPECIFIC_FIELDS
    assert FORM_FIELD_REGISTRY.entradas_saidas.search[-1] == EntradaSaidaField.DADOS_DOC
    assert len(FORM_FIELD_REGISTRY.transferencias) == 5
    assert len(FORM_FIELD_REGISTRY.caixas_bancos) == 5
