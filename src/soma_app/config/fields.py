from dataclasses import dataclass
from enum import Enum
from typing import Iterable


class EntradaSaidaField(str, Enum):
    PLANO_CONTA = "PLANO_CONTA"
    CENTRO_CUSTO = "CENTRO_CUSTO"
    DESCRICAO = "DESCRICAO"
    VALOR = "VALOR"
    OBS = "OBS"
    DATA_ENTRADA = "DATA_ENTRADA"
    FORMA_PAGAMENTO_ENTRADA = "FORMA_PAGAMENTO_ENTRADA"
    CAIXA_ENTRADA = "CAIXA_ENTRADA"
    DATA_VENCIMENTO_SAIDA = "DATA_VENCIMENTO_SAIDA"
    FORMA_PAGAMENTO_MODAL = "FORMA_PAGAMENTO_MODAL"
    NUM_DOCUMENTO_MODAL = "NUM_DOCUMENTO_MODAL"
    CAIXA_PAGAMENTO_MODAL = "CAIXA_PAGAMENTO_MODAL"
    DATA_BAIXA = "DATA_BAIXA"
    PESQ_DESCRICAO = "PESQ_DESCRICAO"
    DATA_INI = "DATA_INI"
    DATA_FIM = "DATA_FIM"
    RESULT_DOC = "RESULT_DOC"
    DADOS_DOC = "DADOS_DOC"

def field_name(field: EntradaSaidaField | str) -> str:
    if isinstance(field, Enum):
        return str(field.value)
    return str(field)


def field_names(fields: Iterable[EntradaSaidaField]) -> tuple[str, ...]:
    return tuple(field_name(field) for field in fields)


class TransferenciaField(str, Enum):
    CAIXA_SAIDA = "CAIXA_SAIDA"
    VALOR = "VALOR"
    CAIXA_ENTRADA = "CAIXA_ENTRADA"
    DATA = "DATA"
    DESCRICAO = "DESCRICAO"


class CaixaBancoField(str, Enum):
    CAIXA_DIARIO = "CAIXA_DIARIO"
    CAIXA_BANCO = "CAIXA_BANCO"
    CRIANCAS = "D_CRIANCAS"
    CAFE = "VERBO_CAFE"
    LIVRARIA = "VERBO_SHOP"


@dataclass(frozen=True)
class EntradasSaidasFieldRegistry:
    common: tuple[EntradaSaidaField, ...]
    entrada_specific: tuple[EntradaSaidaField, ...]
    saida_specific: tuple[EntradaSaidaField, ...]
    search: tuple[EntradaSaidaField, ...]


@dataclass(frozen=True)
class FormFieldRegistry:
    entradas_saidas: EntradasSaidasFieldRegistry
    transferencias: tuple[TransferenciaField, ...]
    caixas_bancos: tuple[CaixaBancoField, ...]


FORM_FIELD_REGISTRY = FormFieldRegistry(
    entradas_saidas=EntradasSaidasFieldRegistry(
        common=(
            EntradaSaidaField.PLANO_CONTA,
            EntradaSaidaField.CENTRO_CUSTO,
            EntradaSaidaField.DESCRICAO,
            EntradaSaidaField.VALOR,
            EntradaSaidaField.OBS,
        ),
        entrada_specific=(
            EntradaSaidaField.DATA_ENTRADA,
            EntradaSaidaField.FORMA_PAGAMENTO_ENTRADA,
            EntradaSaidaField.CAIXA_ENTRADA,
        ),
        saida_specific=(
            EntradaSaidaField.DATA_VENCIMENTO_SAIDA,
        ),
        search=(
            EntradaSaidaField.PESQ_DESCRICAO,
            EntradaSaidaField.DATA_INI,
            EntradaSaidaField.DATA_FIM,
            EntradaSaidaField.RESULT_DOC,
            EntradaSaidaField.DADOS_DOC,
        ),
    ),
    transferencias=(
        TransferenciaField.CAIXA_SAIDA,
        TransferenciaField.VALOR,
        TransferenciaField.CAIXA_ENTRADA,
        TransferenciaField.DATA,
        TransferenciaField.DESCRICAO,
    ),
    caixas_bancos=(
        CaixaBancoField.CAIXA_DIARIO,
        CaixaBancoField.CAIXA_BANCO,
        CaixaBancoField.CRIANCAS,
        CaixaBancoField.CAFE,
        CaixaBancoField.LIVRARIA,
    ),
)

ENTRADA_COMMON_FIELDS = FORM_FIELD_REGISTRY.entradas_saidas.common
ENTRADA_SPECIFIC_FIELDS = FORM_FIELD_REGISTRY.entradas_saidas.entrada_specific
SAIDA_SPECIFIC_FIELDS = FORM_FIELD_REGISTRY.entradas_saidas.saida_specific
SEARCH_FIELDS = FORM_FIELD_REGISTRY.entradas_saidas.search
