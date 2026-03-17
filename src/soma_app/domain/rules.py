from __future__ import annotations

from dataclasses import dataclass
from typing import List

from soma_app.domain.models import ContaOrdemRow, LinhaStatus, TipoMovimento


@dataclass(frozen=True)
class ValidationError:
    field: str
    message: str


def should_process(row: ContaOrdemRow, allow_retry_error: bool = False) -> bool:
    """
    Mantém o comportamento atual:
    - processa Pendente e Em processamento
    - NÃO processa Concluída
    - Em erro só se allow_retry_error=True
    """
    if row.status == LinhaStatus.PENDENTE:
        return True
    if row.status == LinhaStatus.EM_PROCESSAMENTO:
        return True
    if row.status == LinhaStatus.EM_ERRO:
        return bool(allow_retry_error)
    return False


def validate_row(row: ContaOrdemRow) -> List[ValidationError]:
    errs: List[ValidationError] = []

    # obrigatórios gerais (base)
    if not row.tipo:
        errs.append(ValidationError("TIPO", "TIPO é obrigatório"))

    if not row.data_mov:
        errs.append(ValidationError("DATA MOV.", "DATA MOV. é obrigatória"))

    if not row.caixa:
        errs.append(ValidationError("CAIXA", "CAIXA é obrigatória"))

    if not row.centro_custo:
        errs.append(ValidationError("CENTRO DE CUSTO", "CENTRO DE CUSTO é obrigatório"))

    if not row.plano_conta:
        errs.append(ValidationError("PLANO DE CONTA", "PLANO DE CONTA é obrigatório"))

    if not row.forma_pagamento:
        errs.append(ValidationError("FORMA DE PAGAMENTO", "FORMA DE PAGAMENTO é obrigatória"))

    if not row.importancia:
        errs.append(ValidationError("IMPORTÂNCIA", "IMPORTÂNCIA é obrigatória"))

    if not row.descricao_soma:
        errs.append(ValidationError("DESCRIÇÃO SOMA", "DESCRIÇÃO SOMA é obrigatória"))

    # obrigatórios específicos
    if row.tipo == TipoMovimento.SAIDA:
        if not row.caixa_saida:
            errs.append(ValidationError("CAIXA SAIDA", "CAIXA SAIDA é obrigatória para Saída"))

    if row.tipo == TipoMovimento.TRANSFERENCIA:
        if not row.caixa_saida:
            errs.append(ValidationError("CAIXA SAIDA", "CAIXA SAIDA é obrigatória para Transferência"))

    return errs