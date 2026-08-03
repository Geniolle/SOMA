#!/usr/bin/env python3
"""Sistema de scoring de confiança aprimorado para OCR."""

from dataclasses import dataclass
from typing import Any, Callable, TYPE_CHECKING

import numpy as np

from ocr_validators import (
    validate_column_coherence,
    validate_row_consistency,
    validate_description_semantics,
)

if TYPE_CHECKING:
    from main import Movement


@dataclass
class ConfidenceBreakdown:
    """Detalhamento do cálculo de confiança."""
    ocr_score: float
    ocr_weight: float
    column_score: float
    column_weight: float
    validation_score: float
    validation_weight: float
    history_score: float
    history_weight: float
    final_score: float
    factors: dict[str, Any]


def cross_validate_movement(movement: "Movement", cfg: dict) -> dict[str, bool]:
    """
    Executa múltiplas validações cruzadas:
    - Coerência data/valor
    - Consistência de montante
    - Validação de negócio

    Args:
        movement: Movimento a validar
        cfg: Configuração

    Returns:
        Dict com resultados de cada validação
    """
    from main import parse_money, extract_day_month, valid_date

    validations = {}

    # 1. Validação de Coerência Data
    d_mov = extract_day_month(movement.data_movimento)
    d_val = extract_day_month(movement.data_valor)

    if d_mov and d_val:
        # Data de valor não deve ser ANTES de data de movimento
        # Exceção: 30/6 -> 1/7 (última dia do mês)
        if d_mov == d_val:
            validations["date_consistency"] = True
        elif d_val < d_mov and not (d_mov == (30, 6) and d_val == (1, 7)):
            validations["date_order"] = False
        else:
            validations["date_order"] = True
    else:
        validations["date_consistency"] = True  # Assume neutro

    # 2. Validação de Consistência de Montante
    debit = parse_money(movement.debito_eur)
    credit = parse_money(movement.credito_eur)

    if debit is not None and credit is not None:
        # Não pode ter débito e crédito simultaneamente
        validations["amount_exclusivity"] = False
    else:
        validations["amount_exclusivity"] = True

    # 3. Validação de Valor Típico
    validations["amount_reasonable"] = True
    if debit is not None:
        # Valores típicos em extratos de cartão
        if debit < 0.01 or debit > 10000:
            validations["amount_reasonable"] = False

    if credit is not None:
        if credit < 0.01 or credit > 10000:
            validations["amount_reasonable"] = False

    # 4. Validação de País (se presente)
    validations["country_valid"] = True
    if movement.pais:
        valid_countries = {"IRL", "USA", "PRT", "ESP", "GBR", "DEU", "FRA", "NLD", "LUX", "ITA", "CHE"}
        if movement.pais.upper() not in valid_countries:
            validations["country_valid"] = False

    # 5. Validação de descrição não é vazia/muito curta
    validations["description_valid"] = len(movement.descricao.strip()) >= 3

    return validations


def calculate_enhanced_confidence_score(
    movement: "Movement",
    row_words: list,
    table_cfg: dict,
    historical_context: dict | None = None,
    cfg: dict | None = None,
    verbose: bool = False
) -> ConfidenceBreakdown:
    """
    Calcula score de confiança em múltiplas dimensões.

    Retorna score 0.0-1.0 onde 1.0 é confiança máxima.

    Args:
        movement: Movimento a avaliar
        row_words: Palavras da linha (para análise espacial)
        table_cfg: Configuração de tabela com colunas
        historical_context: Contexto histórico/padrões conhecidos
        cfg: Configuração geral
        verbose: Se deve incluir detalhes dos cálculos

    Returns:
        ConfidenceBreakdown com scores e detalhes
    """
    scores = {}
    factors = {}

    # 1. Confiança de OCR pura (30%)
    ocr_weight = 0.30
    if row_words:
        ocr_confidence = np.mean([w.confidence for w in row_words])
    else:
        ocr_confidence = movement.confidence

    # Normalizar confiança se muito alta/baixa
    if ocr_confidence > 1.0:
        ocr_confidence = 1.0
    elif ocr_confidence < 0.0:
        ocr_confidence = 0.0

    ocr_score = ocr_confidence * ocr_weight
    scores["ocr"] = ocr_score
    factors["ocr_base_confidence"] = ocr_confidence

    # 2. Coerência de coluna (25%)
    column_weight = 0.25
    image_width = table_cfg.get("width", 1000)
    columns = table_cfg.get("columns", {})

    if row_words and columns:
        column_validity = validate_column_coherence(row_words, columns, image_width)
        col_score_ratio = 1.0 - (len(column_validity) / max(len(columns), 1))
    else:
        col_score_ratio = 1.0  # Assume perfeito se sem dados

    column_score = col_score_ratio * column_weight
    scores["column"] = column_score
    factors["column_issues"] = len(column_validity) if row_words and columns else 0

    # 3. Validação de dados (20%)
    validation_weight = 0.20
    cross_vals = cross_validate_movement(movement, cfg or {})
    validation_pass_rate = sum(cross_vals.values()) / max(len(cross_vals), 1)
    validation_score = validation_pass_rate * validation_weight

    scores["validation"] = validation_score
    factors["validation_checks"] = sum(cross_vals.values())
    factors["total_validations"] = len(cross_vals)
    factors["failed_checks"] = [k for k, v in cross_vals.items() if not v]

    # 4. Descrição semântica (15%)
    semantic_weight = 0.15
    is_semantic_valid, reason = validate_description_semantics(movement.descricao)
    semantic_score = (1.0 if is_semantic_valid else 0.5) * semantic_weight

    scores["semantic"] = semantic_score
    factors["semantic_valid"] = is_semantic_valid
    factors["semantic_reason"] = reason

    # 5. Consistência com histórico (10%)
    history_weight = 0.10
    if historical_context:
        desc_match = historical_context.get("similar_description_found", False)
        pattern_match = historical_context.get("pattern_matches", False)
        merchant_known = historical_context.get("merchant_known", False)

        hist_score_base = (desc_match + pattern_match + merchant_known) / 3
        history_score = hist_score_base * history_weight

        factors["history_desc_match"] = desc_match
        factors["history_pattern_match"] = pattern_match
        factors["history_merchant_known"] = merchant_known
    else:
        history_score = history_weight * 0.5  # Assume neutro sem histórico
        factors["history_available"] = False

    scores["history"] = history_score

    # Score final ponderado
    final_score = sum(scores.values())
    final_score = min(max(final_score, 0.0), 1.0)  # Clamp 0-1

    if verbose:
        factors["score_breakdown"] = {
            "ocr": f"{ocr_score:.3f} ({ocr_weight*100:.0f}%)",
            "column": f"{column_score:.3f} ({column_weight*100:.0f}%)",
            "validation": f"{validation_score:.3f} ({validation_weight*100:.0f}%)",
            "semantic": f"{semantic_score:.3f} ({semantic_weight*100:.0f}%)",
            "history": f"{history_score:.3f} ({history_weight*100:.0f}%)",
        }

    return ConfidenceBreakdown(
        ocr_score=ocr_score,
        ocr_weight=ocr_weight,
        column_score=column_score,
        column_weight=column_weight,
        validation_score=validation_score,
        validation_weight=validation_weight,
        history_score=history_score,
        history_weight=history_weight,
        final_score=final_score,
        factors=factors,
    )


def detect_potential_false_rejection(
    movement: "Movement",
    cfg: dict,
    trusted_descriptions: set[str],
    merchant_patterns: dict | None = None
) -> tuple[bool, str]:
    """
    Detecta se um movimento foi rejeitado incorretamente.

    Args:
        movement: Movimento rejeitado
        cfg: Configuração
        trusted_descriptions: Descrições confiáveis
        merchant_patterns: Padrões históricos de comerciantes

    Returns:
        (deve_reconsiderar, motivo)
    """
    from main import valid_date

    if movement.status != "REVISÃO":
        return False, "not_rejected"

    # Se a descrição está na lista confiável, reconsidera
    if movement.descricao.upper().strip() in trusted_descriptions:
        return True, "in_trusted_list"

    # Se foi rejeitado apenas por confiança OCR baixa
    if "confiança OCR" in movement.motivos_revisao:
        ocr_confidence = movement.confidence

        # Mas a descrição é legível e o montante faz sentido
        if ocr_confidence > 0.65 and len(movement.descricao) > 5:
            # Reconsidere se o padrão é conhecido
            if merchant_patterns:
                desc_upper = movement.descricao.upper().strip()
                if desc_upper in merchant_patterns:
                    return True, "merchant_known_pattern"

    # Se tem data válida mas foi rejeitado por outra razão
    validation_cfg = cfg.get("validation", {})
    if valid_date(movement.data_movimento, validation_cfg):
        reasons = movement.motivos_revisao.split(";")
        reason_count = len([r for r in reasons if r.strip()])

        if reason_count == 1 and "confiança" not in movement.motivos_revisao:
            # Único motivo e não é confiança - avaliar se é crítico
            if movement.debito_eur or movement.credito_eur:
                # Tem montante e data válida, talvez seja válido
                return True, "has_valid_date_and_amount"

    # Se tem múltiplas validações de negócio passando
    cross_vals = cross_validate_movement(movement, cfg)
    passing_validations = sum(cross_vals.values())

    if passing_validations >= len(cross_vals) - 1:  # Quase tudo passa
        return True, "almost_all_validations_pass"

    return False, "rejection_seems_justified"


def calculate_rejection_confidence_threshold(
    movements: list["Movement"],
    target_accuracy: float = 0.95
) -> float:
    """
    Calcula um threshold de confiança dinâmico baseado nos dados.

    Args:
        movements: Lista de movimentos processados
        target_accuracy: Acurácia alvo (ex: 0.95 = 95%)

    Returns:
        Score de confiança recomendado como threshold
    """
    if not movements:
        return 0.75

    valid_scores = [m.confidence for m in movements if m.status == "VÁLIDO"]
    review_scores = [m.confidence for m in movements if m.status == "REVISÃO"]

    if not valid_scores or not review_scores:
        return 0.75

    # Encontrar threshold que maximize acurácia
    threshold_candidates = np.linspace(0.5, 0.95, 20)
    best_threshold = 0.75
    best_accuracy = 0.0

    for threshold in threshold_candidates:
        # Simular decisão
        correct = (
            sum(1 for s in valid_scores if s >= threshold) +
            sum(1 for s in review_scores if s < threshold)
        )
        accuracy = correct / len(movements)

        if accuracy >= target_accuracy and accuracy > best_accuracy:
            best_accuracy = accuracy
            best_threshold = threshold

    return best_threshold
