#!/usr/bin/env python3
"""Pós-processador OCR: filtro e correcção de linhas de cabeçalho e padrões conhecidos."""

import re
from typing import Optional


# Padrões de cabeçalho que aparecem na OCR
HEADER_KEYWORDS = {
    "taxa", "débita", "crédita", "original", "câmbio", "eur",
    "valor", "data", "descrição", "movimento", "comissão taxa",
    "taxa cambio", "débito", "crédito", "moeda"
}

# Palavras de cabeçalho que indicam 100% que é cabeçalho
HEADER_PATTERNS = [
    r"taxa\s+d[eé]bita",
    r"original\s+taxa",
    r"débita\s+eur",
    r"câmbio\s+eur",
    r"taxa\s+câmbio",
]

# Valores errados que parecem cabeçalho
WRONG_VALUES = {
    "taxa": "",
    "débita": "",
    "crédita": "",
    "débito": "",
    "crédito": "",
    "eur": "",
    "eur ( )": "",
    "câmbio": "",
    "original": "",
}


def is_header_line(texto_ocr: str, fields: dict) -> bool:
    """
    Detecta se a linha é um cabeçalho.

    Retorna True se:
    - Contém padrões de cabeçalho
    - Valores monetários contêm palavras de cabeçalho
    - Descrição está vazia E valores estão errados
    """
    if not texto_ocr:
        return False

    texto_lower = texto_ocr.lower().strip()

    # Verificar padrões explícitos de cabeçalho
    for pattern in HEADER_PATTERNS:
        if re.search(pattern, texto_lower):
            return True

    # Se debito_eur ou credito_eur contêm palavras de cabeçalho, é cabeçalho
    debito_lower = fields.get("debito_eur", "").lower().strip()
    credito_lower = fields.get("credito_eur", "").lower().strip()

    if debito_lower in ["taxa", "débita", "débito", "crédita", "crédito", "câmbio", "eur", "eur ( )"]:
        return True
    if credito_lower in ["taxa", "débita", "débito", "crédita", "crédito", "câmbio", "eur", "eur ( )"]:
        return True

    # Se taxa_cambio é "Original" (cabeçalho), é cabeçalho
    taxa_lower = fields.get("taxa_cambio", "").lower().strip()
    if taxa_lower == "original":
        return True

    # Se descrição vazia + ambos os valores são palavras de cabeçalho = cabeçalho
    descricao = fields.get("descricao", "").strip()
    if not descricao or len(descricao) < 3:
        if debito_lower in WRONG_VALUES and credito_lower in WRONG_VALUES:
            if debito_lower or credito_lower:  # Pelo menos um está preenchido errado
                return True

    return False


def clean_field_value(value: str) -> str:
    """Remove espaços extras e normaliza valor."""
    if not value:
        return ""

    value = value.strip()

    # Se o valor é uma palavra de cabeçalho, retorna vazio
    if value.lower() in WRONG_VALUES:
        return ""

    return value


def correct_known_patterns(fields: dict) -> tuple[dict, list[str]]:
    """
    Corrige padrões conhecidos de erro na OCR.

    Retorna:
    - dicionário com campos corrigidos
    - lista com razões de correção
    """
    reasons = []
    corrected = fields.copy()

    # Corrigir valores que parecem cabeçalho
    for field in ["debito_eur", "credito_eur", "taxa_cambio"]:
        value = corrected.get(field, "").strip()
        if value.lower() in WRONG_VALUES:
            corrected[field] = ""
            reasons.append(f"{field} corrigido (era palavra-chave: '{value}')")

    # Se ambos débito e crédito estão vazios, marcar para revisão
    debito = corrected.get("debito_eur", "").strip()
    credito = corrected.get("credito_eur", "").strip()
    if not debito and not credito:
        reasons.append("débito/crédito ausente (ambos vazios)")

    # Descrição muito curta
    descricao = corrected.get("descricao", "").strip()
    if descricao and len(descricao) < 3:
        reasons.append("descrição muito curta")
    elif not descricao:
        reasons.append("descrição ausente")

    # Limpar espaços extras em todos os campos
    for field in corrected:
        if isinstance(corrected[field], str):
            corrected[field] = clean_field_value(corrected[field])

    return corrected, reasons


def should_mark_for_review(texto_ocr: str, fields: dict, reasons: list) -> tuple[bool, list[str]]:
    """
    Determina se a linha deve ser marcada para REVISÃO.

    Retorna:
    - bool: True se deve revisar
    - list: razões adicionais de revisão
    """
    review_reasons = reasons.copy()

    # Se é cabeçalho, marcar para revisão
    if is_header_line(texto_ocr, fields):
        review_reasons.append("linha de cabeçalho detectada")
        return True, review_reasons

    # Se tem razões de correção, marcar para revisão
    if reasons:
        return True, review_reasons

    return False, review_reasons


def postprocess_ocr_line(texto_ocr: str, fields: dict) -> tuple[dict, str, list[str]]:
    """
    Pós-processa uma linha OCR.

    Retorna:
    - campos corrigidos
    - status (VÁLIDO ou REVISÃO)
    - razões de revisão
    """
    # Primeiro, corrigir padrões conhecidos
    corrected_fields, correction_reasons = correct_known_patterns(fields)

    # Depois, verificar se deve revisar
    should_review, all_reasons = should_mark_for_review(texto_ocr, corrected_fields, correction_reasons)

    status = "REVISÃO" if should_review else "VÁLIDO"

    return corrected_fields, status, all_reasons
