from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Iterable, List, Optional

import unicodedata


def _strip(s: Any) -> str:
    return ("" if s is None else str(s)).strip()


def _norm_basic(s: str) -> str:
    """
    Normalização para comparação:
    - strip
    - lowercase
    - remove acentos
    - colapsa espaços
    """
    s = _strip(s).lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())
    return s


def _get_any(raw: Dict[str, Any], keys: Iterable[str], default: str = "") -> str:
    for k in keys:
        if k in raw:
            return _strip(raw.get(k))
    return default


class TipoMovimento(str, Enum):
    ENTRADA = "Entrada"
    SAIDA = "Saída"
    TRANSFERENCIA = "Transferência"

    @staticmethod
    def from_sheet_value(value: Any) -> "TipoMovimento":
        v = _norm_basic(_strip(value))
        if v in ("entrada", "ent", "in"):
            return TipoMovimento.ENTRADA
        if v in ("saida", "saída", "out"):
            return TipoMovimento.SAIDA
        if v in ("transferencia", "transferência", "transf", "trx"):
            return TipoMovimento.TRANSFERENCIA
        raise ValueError(f"TIPO inválido: {value!r}")


class LinhaStatus(str, Enum):
    PENDENTE = "Pendente"
    EM_PROCESSAMENTO = "Em processamento"
    EM_ERRO = "Em erro"
    CONCLUIDA = "Concluída"


class WorkflowKind(str, Enum):
    ENTRADAS_SAIDAS = "Entradas/Saídas"
    TRANSFERENCIAS = "Transferências"


def status_from_doc_soma(doc_soma: Any) -> LinhaStatus:
    v = _strip(doc_soma)
    if not v:
        return LinhaStatus.PENDENTE

    vn = _norm_basic(v)
    if vn == "em processamento":
        return LinhaStatus.EM_PROCESSAMENTO
    if vn == "em erro":
        return LinhaStatus.EM_ERRO

    # Se tem qualquer "conteúdo" diferente desses estados,
    # consideramos concluída (normalmente é um número/doc).
    return LinhaStatus.CONCLUIDA


def workflow_for_tipo(tipo: TipoMovimento) -> WorkflowKind:
    if tipo == TipoMovimento.TRANSFERENCIA:
        return WorkflowKind.TRANSFERENCIAS
    return WorkflowKind.ENTRADAS_SAIDAS


@dataclass(frozen=True)
class ContaOrdemRow:
    row_number: int
    tipo: TipoMovimento

    data_mov: str
    caixa: str
    caixa_saida: str
    centro_custo: str
    plano_conta: str
    forma_pagamento: str
    importancia: str
    descricao_soma: str

    doc_soma: str
    dados_doc: str
    iduser: str
    timestamp: str

    raw: Dict[str, Any]

    @property
    def status(self) -> LinhaStatus:
        return status_from_doc_soma(self.doc_soma)

    @property
    def workflow(self) -> WorkflowKind:
        return workflow_for_tipo(self.tipo)

    @staticmethod
    def from_table_row(row_number: int, raw: Dict[str, Any]) -> "ContaOrdemRow":
        # Headers observados / variantes comuns
        tipo = TipoMovimento.from_sheet_value(_get_any(raw, ["TIPO"], default=""))

        data_mov = _get_any(raw, ["DATA MOV.", "DATA MOV", "DATA_MOV", "DATA"], default="")
        caixa = _get_any(raw, ["CAIXA"], default="")
        caixa_saida = _get_any(raw, ["CAIXA SAIDA", "CAIXA SAÍDA", "CAIXA_SAIDA"], default="")
        centro_custo = _get_any(raw, ["CENTRO DE CUSTO", "CENTRO CUSTO", "CC"], default="")
        plano_conta = _get_any(raw, ["PLANO DE CONTA", "PLANO DE CONTAS", "CONTA"], default="")
        forma_pagamento = _get_any(raw, ["FORMA DE PAGAMENTO", "FORMA PAGAMENTO"], default="")
        importancia = _get_any(raw, ["IMPORTÂNCIA", "IMPORTANCIA"], default="")
        descricao_soma = _get_any(raw, ["DESCRIÇÃO SOMA", "DESCRICAO SOMA", "DESCRIÇÃO", "DESCRICAO"], default="")

        doc_soma = _get_any(raw, ["DOC. SOMA", "DOC SOMA"], default="")
        dados_doc = _get_any(raw, ["DADOS DOC", "DADOS_DOCTO", "DADOS"], default="")
        iduser = _get_any(raw, ["IDUSER", "USER", "USUARIO"], default="")
        timestamp = _get_any(raw, ["TIMESTAMP", "DATAHORA", "DATA HORA"], default="")

        return ContaOrdemRow(
            row_number=row_number,
            tipo=tipo,
            data_mov=data_mov,
            caixa=caixa,
            caixa_saida=caixa_saida,
            centro_custo=centro_custo,
            plano_conta=plano_conta,
            forma_pagamento=forma_pagamento,
            importancia=importancia,
            descricao_soma=descricao_soma,
            doc_soma=doc_soma,
            dados_doc=dados_doc,
            iduser=iduser,
            timestamp=timestamp,
            raw=dict(raw),
        )


@dataclass(frozen=True)
class CaixaRow:
    row_number: int
    caixa: str
    caixa_diario: str
    caixa_banco: str
    raw: Dict[str, Any]

    @staticmethod
    def from_table_row(row_number: int, raw: Dict[str, Any]) -> "CaixaRow":
        caixa = _get_any(raw, ["CAIXA"], default="")
        caixa_diario = _get_any(raw, ["CAIXA DIÁRIO", "CAIXA DIARIO"], default="")
        caixa_banco = _get_any(raw, ["CAIXA BANCO"], default="")
        return CaixaRow(
            row_number=row_number,
            caixa=caixa,
            caixa_diario=caixa_diario,
            caixa_banco=caixa_banco,
            raw=dict(raw),
        )