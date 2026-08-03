#!/usr/bin/env python3
"""Métricas e análise de qualidade OCR."""

from collections import Counter
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Any, TYPE_CHECKING

import numpy as np

if TYPE_CHECKING:
    from main import Movement


@dataclass
class OCRQualityMetrics:
    """Métricas de qualidade agregadas."""
    timestamp: str
    total_movements: int
    valid_count: int
    review_count: int
    success_rate: float
    avg_confidence: float
    median_confidence: float
    min_confidence: float
    max_confidence: float
    confidence_std: float
    rejection_reasons: dict[str, int]
    top_rejection_reasons: list[tuple[str, int]]
    quality_trend: str
    metrics_version: str = "1.0"


def generate_ocr_quality_metrics(movements: list["Movement"], cfg: dict | None = None) -> OCRQualityMetrics:
    """
    Gera métricas de qualidade do OCR para monitoramento.

    Args:
        movements: Lista de movimentos processados

    Returns:
        Métricas agregadas
    """
    total = len(movements)

    if total == 0:
        return OCRQualityMetrics(
            timestamp=datetime.now().isoformat(),
            total_movements=0,
            valid_count=0,
            review_count=0,
            success_rate=0.0,
            avg_confidence=0.0,
            median_confidence=0.0,
            min_confidence=0.0,
            max_confidence=0.0,
            confidence_std=0.0,
            rejection_reasons={},
            top_rejection_reasons=[],
            quality_trend="unknown"
        )

    validos = sum(1 for m in movements if m.status == "VÁLIDO")
    revisao = sum(1 for m in movements if m.status == "REVISÃO")

    confidence_scores = [m.confidence for m in movements]
    avg_confidence = np.mean(confidence_scores)
    median_confidence = np.median(confidence_scores)
    min_confidence = np.min(confidence_scores)
    max_confidence = np.max(confidence_scores)
    confidence_std = np.std(confidence_scores)

    # Análise de motivos de rejeição
    rejection_reasons = Counter()
    for m in movements:
        if m.status == "REVISÃO" and m.motivos_revisao:
            for reason in m.motivos_revisao.split(";"):
                reason = reason.strip()
                if reason:
                    rejection_reasons[reason] += 1

    top_reasons = rejection_reasons.most_common(5)

    # Determinar tendência de qualidade
    if avg_confidence > 0.90:
        quality_trend = "excellent"
    elif avg_confidence > 0.85:
        quality_trend = "good"
    elif avg_confidence > 0.75:
        quality_trend = "acceptable"
    elif avg_confidence > 0.65:
        quality_trend = "needs_improvement"
    else:
        quality_trend = "poor"

    return OCRQualityMetrics(
        timestamp=datetime.now().isoformat(),
        total_movements=total,
        valid_count=validos,
        review_count=revisao,
        success_rate=(validos / total * 100) if total > 0 else 0.0,
        avg_confidence=avg_confidence,
        median_confidence=median_confidence,
        min_confidence=min_confidence,
        max_confidence=max_confidence,
        confidence_std=confidence_std,
        rejection_reasons=dict(rejection_reasons),
        top_rejection_reasons=top_reasons,
        quality_trend=quality_trend
    )


def compare_metrics(
    metrics1: OCRQualityMetrics,
    metrics2: OCRQualityMetrics
) -> dict[str, Any]:
    """
    Compara dois conjuntos de métricas para detectar tendências.

    Args:
        metrics1: Métricas antigas (linha de base)
        metrics2: Métricas novas

    Returns:
        Dicionário com comparações
    """
    return {
        "timestamp_1": metrics1.timestamp,
        "timestamp_2": metrics2.timestamp,
        "success_rate_change": metrics2.success_rate - metrics1.success_rate,
        "success_rate_change_pct": (
            ((metrics2.success_rate - metrics1.success_rate) / metrics1.success_rate * 100)
            if metrics1.success_rate > 0 else 0.0
        ),
        "avg_confidence_change": metrics2.avg_confidence - metrics1.avg_confidence,
        "total_movements_change": metrics2.total_movements - metrics1.total_movements,
        "quality_improved": metrics2.quality_trend >= metrics1.quality_trend,
        "trend_direction": "improving" if metrics2.quality_trend >= metrics1.quality_trend else "declining",
    }


def format_metrics_report(metrics: OCRQualityMetrics) -> str:
    """
    Formata métricas como relatório legível em texto.

    Args:
        metrics: Métricas a formatar

    Returns:
        String com relatório formatado
    """
    report = f"""
╔════════════════════════════════════════════════════════════════╗
║              RELATÓRIO DE QUALIDADE OCR                        ║
╚════════════════════════════════════════════════════════════════╝

📊 Resumo Geral
  • Timestamp: {metrics.timestamp}
  • Total de movimentos: {metrics.total_movements}
  • Movimentos válidos: {metrics.valid_count} ({metrics.success_rate:.1f}%)
  • Movimentos para revisão: {metrics.review_count}

📈 Confiança
  • Média: {metrics.avg_confidence:.1%}
  • Mediana: {metrics.median_confidence:.1%}
  • Intervalo: {metrics.min_confidence:.1%} - {metrics.max_confidence:.1%}
  • Desvio padrão: {metrics.confidence_std:.1%}

🎯 Qualidade Geral
  • Status: {metrics.quality_trend.upper()}
  • Taxa de sucesso alvo: ≥95%
  • Diferença do alvo: {95 - metrics.success_rate:.1f}%

❌ Top Motivos de Rejeição
"""

    for reason, count in metrics.top_rejection_reasons[:5]:
        pct = (count / metrics.review_count * 100) if metrics.review_count > 0 else 0
        report += f"  • {reason}: {count} ({pct:.0f}% das rejeições)\n"

    report += f"\n{'═' * 64}\n"

    return report


def get_metrics_summary(metrics: OCRQualityMetrics) -> dict[str, Any]:
    """
    Retorna um resumo executivo das métricas.

    Args:
        metrics: Métricas a resumir

    Returns:
        Dicionário com informações-chave
    """
    return {
        "success_rate": f"{metrics.success_rate:.1f}%",
        "quality_status": metrics.quality_trend,
        "avg_confidence": f"{metrics.avg_confidence:.1%}",
        "valid_movements": metrics.valid_count,
        "movements_for_review": metrics.review_count,
        "primary_issue": (
            metrics.top_rejection_reasons[0][0]
            if metrics.top_rejection_reasons
            else "none"
        ),
        "recommendation": get_recommendation(metrics),
    }


def get_recommendation(metrics: OCRQualityMetrics) -> str:
    """
    Gera recomendação baseada nas métricas.

    Args:
        metrics: Métricas a analisar

    Returns:
        Recomendação em texto
    """
    if metrics.success_rate >= 95:
        return "✅ Sistema operando dentro dos padrões. Continuar monitorando."

    if metrics.success_rate >= 85:
        return "⚠️ Sistema aceitável mas com espaço para melhoria. Revisar padrões de rejeição."

    if metrics.confidence_std > 0.20:
        return "⚠️ Alta variabilidade de confiança. Pode indicar problemas de qualidade de imagem."

    if "confiança OCR baixa" in dict(metrics.rejection_reasons):
        return "🔧 Melhorar pré-processamento de imagem ou aumentar threshold de confiança."

    if "data movimento inválida" in dict(metrics.rejection_reasons):
        return "🔧 Revisar formatação de datas e padrões de normalização."

    return "📊 Revisar todos os motivos de rejeição para identificar padrões."


def identify_problematic_merchants(
    movements: list["Movement"],
    min_occurrences: int = 2
) -> dict[str, dict]:
    """
    Identifica comerciantes que frequentemente resultam em rejeições.

    Args:
        movements: Lista de movimentos
        min_occurrences: Número mínimo de ocorrências para considerar

    Returns:
        Dict com {descrição: {stats}}
    """
    merchant_stats = {}

    for m in movements:
        desc = m.descricao.upper().strip()

        if desc not in merchant_stats:
            merchant_stats[desc] = {
                "total": 0,
                "valid": 0,
                "review": 0,
                "avg_confidence": [],
                "rejection_reasons": [],
            }

        stats = merchant_stats[desc]
        stats["total"] += 1
        stats["avg_confidence"].append(m.confidence)

        if m.status == "VÁLIDO":
            stats["valid"] += 1
        else:
            stats["review"] += 1
            if m.motivos_revisao:
                stats["rejection_reasons"].append(m.motivos_revisao)

    # Filtrar e calcular stats
    problematic = {}
    for desc, stats in merchant_stats.items():
        if stats["total"] >= min_occurrences and stats["review"] > 0:
            rejection_rate = stats["review"] / stats["total"]
            if rejection_rate > 0.33:  # > 33% de rejeição
                problematic[desc] = {
                    "total_occurrences": stats["total"],
                    "valid": stats["valid"],
                    "rejected": stats["review"],
                    "rejection_rate": rejection_rate,
                    "avg_confidence": np.mean(stats["avg_confidence"]),
                    "common_issues": Counter(stats["rejection_reasons"]).most_common(2),
                }

    return dict(sorted(
        problematic.items(),
        key=lambda x: x[1]["rejection_rate"],
        reverse=True
    ))


def export_metrics_json(metrics: OCRQualityMetrics) -> dict:
    """Exporta métricas como dicionário JSON."""
    data = asdict(metrics)
    # Converter Counter para dict se necessário
    if isinstance(data.get("rejection_reasons"), Counter):
        data["rejection_reasons"] = dict(data["rejection_reasons"])
    return data
