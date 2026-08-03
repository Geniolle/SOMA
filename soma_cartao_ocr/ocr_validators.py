#!/usr/bin/env python3
"""Validadores avançados para OCR de extratos de cartão."""

import re
from dataclasses import dataclass
from typing import Any

import numpy as np


@dataclass
class Word:
    """Representação de uma palavra detectada. (copiado de main.py para compatibilidade)"""
    text: str
    x0: int
    y0: int
    x1: int
    y1: int
    confidence: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


# Correções automáticas comuns de OCR
COMMON_OCR_FIXES = {
    # Padrões de data: espaço ou ponto em vez de barra
    r"(\d{2})\s+(\d{2})$": r"\1/\2",
    r"(\d{2})[.,](\d{2})$": r"\1/\2",

    # Erros de caracteres confundidos em nomes de negócios
    "COMISSAD": "COMISSÃO",
    "COMISSAQ": "COMISSÃO",
    "COMIÇAO": "COMISSÃO",
    "COMISSAO": "COMISSÃO",
    "MERCADORÍA": "MERCADONA",
    "MERCADOÑA": "MERCADONA",
    "SUPERMERCAD": "SUPERMERCADO",
    "FARMACÍA": "FARMÁCIA",
    "FARMAÇIA": "FARMÁCIA",
    "FARMACIA": "FARMÁCIA",
    "RECHEIO": "RECHEIO",  # Manter como está
    "RECHEI0": "RECHEIO",
    "RECHETO": "RECHEIO",
    "CARRY BRAG": "CARRY BRAGA",
    "CARRYBRAGA": "CARRY BRAGA",
    "CANTA": "CANVA",
    "CANVA": "CANVA",
    "CANVA": "CANVA",
    "LEVANT NUMERARIO": "LEVANT. NUMERÁRIO",
    "LEVANTUMERARIO": "LEVANT. NUMERÁRIO",
}


def apply_ocr_corrections(text: str, confidence: float, auto_correct: bool = True) -> tuple[str, bool]:
    """
    Aplica correções automáticas baseadas em confiança e padrões conhecidos.

    Args:
        text: Texto a corrigir
        confidence: Nível de confiança do OCR (0.0-1.0)
        auto_correct: Se deve aplicar correções automáticas

    Returns:
        (texto_corrigido, foi_corrigido)
    """
    if not auto_correct or not text:
        return text, False

    original = text
    corrected = text

    for pattern, replacement in COMMON_OCR_FIXES.items():
        if pattern.startswith("("):  # Regex pattern
            corrected = re.sub(pattern, replacement, corrected, flags=re.IGNORECASE)
        else:
            # Substituição de string simples - case-insensitive
            if pattern.lower() in corrected.lower():
                # Preservar case original
                corrected = re.sub(
                    re.escape(pattern),
                    replacement,
                    corrected,
                    flags=re.IGNORECASE
                )

    was_corrected = original != corrected
    return corrected, was_corrected


def validate_column_coherence(
    row: list[Word],
    column_bounds: dict[str, tuple[float, float]],
    image_width: int
) -> dict[str, dict[str, Any]]:
    """
    Valida se as palavras de cada coluna fazem sentido semântico.
    Detecta quando palavras sangram incorretamente entre colunas.

    Args:
        row: Lista de palavras na linha
        column_bounds: Dict com (nome_coluna, (x_min_ratio, x_max_ratio))
        image_width: Largura da imagem em pixels

    Returns:
        Dict com {nome_coluna: {problemas}}
    """
    issues = {}

    for col_name, (x_min_ratio, x_max_ratio) in column_bounds.items():
        x_min = image_width * x_min_ratio
        x_max = image_width * x_max_ratio
        words_in_col = [w for w in row if x_min <= w.cx < x_max]

        if not words_in_col:
            issues[col_name] = {"reason": "no_words", "severity": "low"}
            continue

        col_info = {
            "word_count": len(words_in_col),
            "words": [w.text for w in words_in_col],
        }

        # 1. Verificar coerência de confiança
        confidence_scores = [w.confidence for w in words_in_col]
        avg_confidence = np.mean(confidence_scores)
        col_info["avg_confidence"] = avg_confidence

        if avg_confidence < 0.60:
            col_info["issue"] = "low_column_confidence"
            col_info["severity"] = "medium"
            issues[col_name] = col_info
            continue

        # 2. Detectar inconsistência de altura
        heights = [w.y1 - w.y0 for w in words_in_col]
        if len(set(heights)) > 1:
            height_variance = max(heights) - min(heights)
            if height_variance > min(heights) * 0.3:  # > 30% de variação
                col_info["issue"] = "height_inconsistency"
                col_info["height_variance"] = height_variance
                col_info["severity"] = "low"
                issues[col_name] = col_info
                continue

        # 3. Detectar espaços anormais entre palavras
        if len(words_in_col) > 1:
            sorted_words = sorted(words_in_col, key=lambda w: w.x0)
            gaps = [sorted_words[i+1].x0 - sorted_words[i].x1
                   for i in range(len(sorted_words)-1)]
            if gaps:
                median_gap = np.median(gaps)
                max_gap = max(gaps)
                if median_gap > 0 and max_gap > 3 * median_gap:
                    col_info["issue"] = "unusual_spacing"
                    col_info["gap_ratio"] = max_gap / median_gap
                    col_info["severity"] = "low"
                    issues[col_name] = col_info

    return issues


def validate_row_consistency(row: list[Word], min_confidence: float = 0.5) -> tuple[bool, str, dict]:
    """
    Valida se uma linha extraída é realmente uma linha de dados
    ou apenas ruído/artefatos.

    Args:
        row: Lista de palavras na linha
        min_confidence: Confiança mínima esperada

    Returns:
        (é_válida, motivo, detalhes)
    """
    details = {
        "word_count": len(row),
        "has_date_pattern": False,
        "avg_confidence": 0.0,
        "unique_words": 0,
        "y_variance": 0.0,
    }

    if not row:
        return False, "empty_row", details

    # 1. Deve ter pelo menos uma data
    date_pattern = re.compile(r"\d{1,2}[/.\-]\d{1,2}|\d{4}-\d{2}-\d{2}")
    date_words = [w for w in row if date_pattern.search(w.text)]
    details["has_date_pattern"] = len(date_words) > 0

    if not date_words:
        return False, "no_date_pattern", details

    # 2. Deve ter confiança média aceitável
    confidence_scores = [w.confidence for w in row]
    avg_confidence = np.mean(confidence_scores)
    details["avg_confidence"] = avg_confidence

    if avg_confidence < min_confidence:
        return False, f"low_avg_confidence_{avg_confidence:.2f}", details

    # 3. Não deve ser repetição excessiva de palavras
    text_words = [w.text.upper() for w in row]
    unique_words = len(set(text_words))
    details["unique_words"] = unique_words

    if unique_words < len(row) * 0.3 and len(row) > 3:  # Menos de 30% unique
        return False, "too_repetitive", details

    # 4. Deve ter mínimo de palavras significativas
    significant_words = [w for w in row if len(w.text) > 2]
    if len(significant_words) < 2:
        return False, "insufficient_content", details

    # 5. Alinhamento Y deve ser coerente
    y_positions = [w.cy for w in row]
    y_variance = np.var(y_positions) if len(y_positions) > 1 else 0
    details["y_variance"] = y_variance

    avg_word_height = np.mean([w.y1 - w.y0 for w in row])
    y_range = max(y_positions) - min(y_positions)

    if avg_word_height > 0 and y_range > avg_word_height * 0.8:
        # Y spread é muito grande - pode ser linha fantasma
        return False, f"inconsistent_y_alignment_{y_range:.1f}", details

    return True, "valid", details


def validate_description_semantics(desc: str, min_length: int = 3) -> tuple[bool, str]:
    """
    Valida se a descrição faz sentido semântico.
    Detecta quando é apenas ruído/números.

    Args:
        desc: Descrição a validar
        min_length: Comprimento mínimo esperado

    Returns:
        (é_válida, motivo)
    """
    if not desc or len(desc.strip()) == 0:
        return False, "empty"

    desc = desc.strip()

    if len(desc) < min_length:
        return False, f"too_short_{len(desc)}"

    # Detectar descrições apenas com números/símbolos
    alpha_count = sum(c.isalpha() for c in desc)
    if alpha_count == 0:
        return False, "no_alpha_chars"

    alpha_ratio = alpha_count / len(desc)
    if alpha_ratio < 0.3:
        return False, f"mostly_numbers_{alpha_ratio:.1%}"

    # Detectar repetição excessiva de um caractere
    if len(set(desc)) < len(desc) * 0.2 and len(desc) > 5:
        return False, "excessive_repetition"

    # Validar que não é apenas espaços e símbolos
    meaningful_chars = [c for c in desc if c.isalnum()]
    if len(meaningful_chars) < 2:
        return False, "no_meaningful_chars"

    return True, "valid"


def detect_phantom_row(row: list[Word], context_rows: list[list[Word]]) -> tuple[bool, str]:
    """
    Detecta se uma linha é "fantasma" - artefato OCR que não é dados reais.

    Args:
        row: Linha a verificar
        context_rows: Linhas anteriores para contexto

    Returns:
        (é_fantasma, motivo)
    """
    is_consistent, reason, details = validate_row_consistency(row)

    if not is_consistent:
        return True, reason

    # Verificar padrões de "fantasma"

    # 1. Linha é uma repetição exata de linha anterior
    if context_rows:
        prev_row = context_rows[-1]
        prev_text = " ".join(w.text for w in prev_row)
        curr_text = " ".join(w.text for w in row)
        if prev_text == curr_text:
            return True, "exact_duplicate"

    # 2. Linha contém só caracteres de "lixo"
    garbage_chars = set("!@#$%^&*()[]{}\\|;:,<>?~`")
    text = " ".join(w.text for w in row)
    garbage_ratio = sum(c in garbage_chars for c in text) / max(len(text), 1)
    if garbage_ratio > 0.3:
        return True, "high_garbage_ratio"

    # 3. Linha muito próxima (Y) de linha anterior - provável duplicação
    if context_rows:
        prev_row = context_rows[-1]
        prev_y = np.mean([w.cy for w in prev_row])
        curr_y = np.mean([w.cy for w in row])
        if abs(curr_y - prev_y) < 10:  # Menos de 10 pixels
            return True, "too_close_to_previous"

    return False, "likely_valid"
