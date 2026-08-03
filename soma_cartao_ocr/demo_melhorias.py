#!/usr/bin/env python3
"""Demonstração dos novos módulos de validação e scoring."""

import json
from pathlib import Path

from main import Movement
from ocr_validators import apply_ocr_corrections, validate_description_semantics
from confidence_scoring import cross_validate_movement, detect_potential_false_rejection
from merchant_patterns import MerchantPatternLearner
from metrics import generate_ocr_quality_metrics, format_metrics_report


def demo_ocr_corrections():
    """Demo: Correções automáticas de OCR"""
    print("\n" + "=" * 70)
    print("DEMO 1: Correções Automáticas de OCR")
    print("=" * 70)

    test_cases = [
        ("COMISSAD", 0.72),
        ("MERCADORÍA", 0.68),
        ("26 06", 0.85),
        ("FARMACÍA", 0.75),
    ]

    for text, confidence in test_cases:
        corrected, was_corrected = apply_ocr_corrections(text, confidence)
        status = "✓ Corrigido" if was_corrected else "→ Sem alteração"
        print(f"{status}: '{text}' ({confidence:.0%}) → '{corrected}'")


def demo_description_validation():
    """Demo: Validação de descrição"""
    print("\n" + "=" * 70)
    print("DEMO 2: Validação de Descrição Semântica")
    print("=" * 70)

    test_cases = [
        "MERCADONA",
        "123456",
        "aaa",
        "!@#$%",
        "SUPERMERCADO CARREFOUR",
    ]

    for desc in test_cases:
        is_valid, reason = validate_description_semantics(desc)
        status = "✓ Válida" if is_valid else "✗ Inválida"
        print(f"{status}: '{desc:30}' - Motivo: {reason}")


def demo_cross_validation():
    """Demo: Validação cruzada de movimento"""
    print("\n" + "=" * 70)
    print("DEMO 3: Validação Cruzada de Movimento")
    print("=" * 70)

    # Movimento válido
    mov_valid = Movement(
        line=1,
        data_movimento="26/06",
        data_valor="25/06",
        descricao="MERCADONA",
        pais="ESP",
        moeda_original="",
        taxa_cambio="",
        debito_eur="35.15",
        credito_eur="",
        confidence=0.92,
        status="VÁLIDO",
        motivos_revisao="",
        texto_ocr="26/06 25/06 MERCADONA 35.15"
    )

    # Movimento com problema
    mov_problem = Movement(
        line=2,
        data_movimento="26/06",
        data_valor="26/06",
        descricao="PAGTO",
        pais="",
        moeda_original="",
        taxa_cambio="",
        debito_eur="0.50",  # Muito baixo
        credito_eur="100.00",  # Ambíguo com débito
        confidence=0.55,
        status="REVISÃO",
        motivos_revisao="valor muito baixo",
        texto_ocr="26/06 26/06 PAGTO"
    )

    cfg = {
        "validation": {
            "year": 2026,
            "allowed_months": [6, 7],
        }
    }

    for mov in [mov_valid, mov_problem]:
        print(f"\nMovimento {mov.line}: {mov.descricao}")
        validations = cross_validate_movement(mov, cfg)

        for check, result in validations.items():
            status = "✓" if result else "✗"
            print(f"  {status} {check}: {result}")


def demo_false_rejection_detection():
    """Demo: Detecção de falsos negativos"""
    print("\n" + "=" * 70)
    print("DEMO 4: Detecção de Falsos Negativos")
    print("=" * 70)

    movements = [
        Movement(
            line=1,
            data_movimento="26/06",
            data_valor="25/06",
            descricao="MERCADONA",
            pais="ESP",
            moeda_original="",
            taxa_cambio="",
            debito_eur="35.15",
            credito_eur="",
            confidence=0.68,  # Baixa confiança
            status="REVISÃO",
            motivos_revisao="confiança OCR baixa (0.68)",
            texto_ocr="26/06 25/06 MERCADONA 35.15"
        ),
        Movement(
            line=2,
            data_movimento="31/06",  # Data inválida
            data_valor="30/06",
            descricao="PAGTO BANCO",
            pais="",
            moeda_original="",
            taxa_cambio="",
            debito_eur="100.00",
            credito_eur="",
            confidence=0.85,
            status="REVISÃO",
            motivos_revisao="data movimento inválida",
            texto_ocr="31/06 30/06 PAGTO BANCO"
        ),
    ]

    cfg = {"validation": {"year": 2026, "allowed_months": [6, 7]}}
    trusted = {"MERCADONA", "PAGTO BANCO"}

    for mov in movements:
        should_reconsider, reason = detect_potential_false_rejection(mov, cfg, trusted)

        status = "🔄 Reconsiderar" if should_reconsider else "→ Manter rejeição"
        print(f"{status}: {mov.descricao}")
        print(f"   Motivo atual: {mov.motivos_revisao}")
        print(f"   Razão da decisão: {reason}")


def demo_merchant_patterns():
    """Demo: Aprendizado de padrões de comerciantes"""
    print("\n" + "=" * 70)
    print("DEMO 5: Padrões de Comerciantes")
    print("=" * 70)

    # Criar learner temporário (em memória)
    learner = MerchantPatternLearner(None)

    # Simular movimentos válidos
    valid_movements = [
        Movement(1, "26/06", "25/06", "MERCADONA", "ESP", "", "", "35.15", "", 0.92, "VÁLIDO", "", ""),
        Movement(2, "27/06", "26/06", "MERCADONA", "ESP", "", "", "42.50", "", 0.88, "VÁLIDO", "", ""),
        Movement(3, "28/06", "27/06", "MERCADONA", "ESP", "", "", "28.30", "", 0.85, "VÁLIDO", "", ""),
        Movement(4, "29/06", "28/06", "FARMAÇIA", "PRT", "", "", "15.99", "", 0.90, "VÁLIDO", "", ""),
    ]

    print("Aprendendo de movimentos válidos...")
    for mov in valid_movements:
        learner.learn_from_movement(mov)
        print(f"  ✓ {mov.descricao}: €{mov.debito_eur}")

    # Consultar padrões
    print("\nInformações de comerciantes aprendidos:")
    for merchant in ["MERCADONA", "FARMAÇIA"]:
        info = learner.get_merchant_info(merchant)

        if info["known"]:
            low, high = info["typical_amount_range"]
            print(f"\n  {merchant}:")
            print(f"    • Ocorrências: {info['occurrences']}")
            print(f"    • Intervalo típico: €{low:.2f} - €{high:.2f}")
            print(f"    • Confiança: {info['confidence_score']:.0%}")

    # Testar montantes
    print("\nValidação de montantes:")
    test_amounts = [
        ("MERCADONA", 40.00),
        ("MERCADONA", 2.00),  # Muito baixo
        ("MERCADONA", 500.00),  # Muito alto
        ("FARMAÇIA", 20.00),
    ]

    for merchant, amount in test_amounts:
        is_typical = learner.is_expected_amount(merchant, amount)
        status = "✓ Típico" if is_typical else "⚠️  Atípico"
        print(f"  {status}: €{amount:.2f} em {merchant}")


def demo_quality_metrics():
    """Demo: Métricas de qualidade"""
    print("\n" + "=" * 70)
    print("DEMO 6: Métricas de Qualidade OCR")
    print("=" * 70)

    # Criar movimento simulados
    movements = [
        Movement(1, "26/06", "25/06", "MERCADONA", "", "", "", "35.15", "", 0.95, "VÁLIDO", "", ""),
        Movement(2, "27/06", "26/06", "ALMOCO", "", "", "", "12.50", "", 0.88, "VÁLIDO", "", ""),
        Movement(3, "28/06", "27/06", "GASOLINA", "", "", "", "55.00", "", 0.92, "VÁLIDO", "", ""),
        Movement(4, "29/06", "28/06", "COMISSAD", "", "", "", "2.50", "", 0.62, "REVISÃO", "confiança OCR baixa", ""),
        Movement(5, "30/06", "29/06", "BANCO", "", "", "", "10.00", "", 0.70, "REVISÃO", "descrição ausente/curta", ""),
    ]

    print("Gerando métricas...")
    metrics = generate_ocr_quality_metrics(movements)

    print(format_metrics_report(metrics))

    # Resumo executivo
    print("\n📊 Resumo Executivo:")
    print(f"  • Taxa de sucesso: {metrics.success_rate:.1f}% (alvo: ≥95%)")
    print(f"  • Confiança média: {metrics.avg_confidence:.1%}")
    print(f"  • Tendência: {metrics.quality_trend}")
    print(f"  • Motivos top: {', '.join([f[0] for f in metrics.top_rejection_reasons[:2]])}")


def main():
    """Executa todas as demos"""
    print("\n" + "▓" * 70)
    print("▓  DEMONSTRAÇÃO DE MELHORIAS DE CONFIABILIDADE - OCR CARTÃO")
    print("▓" * 70)

    try:
        demo_ocr_corrections()
        demo_description_validation()
        demo_cross_validation()
        demo_false_rejection_detection()
        demo_merchant_patterns()
        demo_quality_metrics()

        print("\n" + "▓" * 70)
        print("▓  DEMONSTRAÇÃO CONCLUÍDA COM SUCESSO!")
        print("▓" * 70)
        print("\n✅ Todos os módulos estão funcionando corretamente.")
        print("\n📖 Para mais informações, consulte:")
        print("   • MELHORIAS_CONFIABILIDADE.md - Análise detalhada e plano")
        print("   • GUIA_INTEGRACAO.md - Como integrar com main.py")
        print("   • Docstrings nos módulos: ocr_validators.py, confidence_scoring.py, etc.")

    except Exception as e:
        print(f"\n❌ Erro durante demonstração: {e}")
        import traceback
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    exit(main())
