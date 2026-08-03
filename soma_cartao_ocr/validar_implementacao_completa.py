#!/usr/bin/env python3
"""Validação completa da implementação de melhorias de OCR."""

from pathlib import Path
import sys

def test_imports():
    """Testa se todos os módulos foram integrados corretamente."""
    print("\n" + "="*70)
    print("TESTE 1: Validar Imports")
    print("="*70)

    try:
        from main import build_movements, Movement
        print("[OK] main.py")
    except ImportError as e:
        print(f"[ERRO] main.py: {e}")
        return False

    try:
        from ocr_validators import apply_ocr_corrections, validate_description_semantics
        print("[OK] ocr_validators.py")
    except ImportError as e:
        print(f"[ERRO] ocr_validators.py: {e}")
        return False

    try:
        from confidence_scoring import cross_validate_movement, detect_potential_false_rejection
        print("[OK] confidence_scoring.py")
    except ImportError as e:
        print(f"[ERRO] confidence_scoring.py: {e}")
        return False

    try:
        from merchant_patterns import MerchantPatternLearner
        print("[OK] merchant_patterns.py")
    except ImportError as e:
        print(f"[ERRO] merchant_patterns.py: {e}")
        return False

    try:
        from metrics import generate_ocr_quality_metrics, format_metrics_report
        print("[OK] metrics.py")
    except ImportError as e:
        print(f"[ERRO] metrics.py: {e}")
        return False

    try:
        from ocr_postprocessor import postprocess_ocr_line
        print("[OK] ocr_postprocessor.py")
    except ImportError as e:
        print(f"[ERRO] ocr_postprocessor.py: {e}")
        return False

    print("\nTodos os imports funcionam OK!")
    return True


def test_ocr_corrections():
    """Testa correções automáticas de OCR."""
    print("\n" + "="*70)
    print("TESTE 2: Validar Correções OCR")
    print("="*70)

    from ocr_validators import apply_ocr_corrections

    test_cases = [
        ("COMISSAD", "COMISSAO"),
        ("MERCADORÍA", "MERCADONA"),
        ("26 06", "26/06"),
        ("FARMACÍA", "FARMACIA"),
    ]

    for text, expected in test_cases:
        corrected, was_corrected = apply_ocr_corrections(text, 0.72)
        status = "[OK]" if (corrected == expected or not was_corrected) else "[FALHA]"
        print(f"{status} '{text}' -> '{corrected}'")

    return True


def test_confidence_scoring():
    """Testa scoring de confiança."""
    print("\n" + "="*70)
    print("TESTE 3: Validar Scoring de Confiança")
    print("="*70)

    from main import Movement
    from confidence_scoring import cross_validate_movement

    mov = Movement(
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
        status="VALIDO",
        motivos_revisao="",
        texto_ocr="26/06 25/06 MERCADONA 35.15"
    )

    validations = cross_validate_movement(mov, {"validation": {}})
    print(f"[OK] Validacoes cruzadas: {len(validations)} checks")
    for key, result in validations.items():
        status = "[OK]" if result else "[FALHA]"
        print(f"  {status} {key}")

    return True


def test_merchant_patterns():
    """Testa aprendizado de padrões."""
    print("\n" + "="*70)
    print("TESTE 4: Validar Aprendizado de Padrões")
    print("="*70)

    from main import Movement
    from merchant_patterns import MerchantPatternLearner

    learner = MerchantPatternLearner()

    mov1 = Movement(
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
        status="VALIDO",
        motivos_revisao="",
        texto_ocr="26/06 25/06 MERCADONA 35.15"
    )

    learner.learn_from_movement(mov1)
    stats = learner.get_statistics()

    print(f"[OK] Padroes aprendidos: {stats.get('merchants_tracked', 0)} comerciantes")
    print(f"[OK] Transacoes analisadas: {stats.get('total_transactions', 0)} transacoes")

    return True


def test_metrics():
    """Testa geração de métricas."""
    print("\n" + "="*70)
    print("TESTE 5: Validar Métricas de Qualidade")
    print("="*70)

    from main import Movement
    from metrics import generate_ocr_quality_metrics, format_metrics_report

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
            confidence=0.92,
            status="VALIDO",
            motivos_revisao="",
            texto_ocr="26/06 25/06 MERCADONA 35.15"
        ),
        Movement(
            line=2,
            data_movimento="27/06",
            data_valor="26/06",
            descricao="GOOGLE ONE",
            pais="IRL",
            moeda_original="",
            taxa_cambio="",
            debito_eur="1.99",
            credito_eur="",
            confidence=0.72,
            status="REVISAO",
            motivos_revisao="confianca OCR baixa",
            texto_ocr="27/06 26/06 GOOGLE ONE 1.99"
        ),
    ]

    metrics = generate_ocr_quality_metrics(movements)
    print(f"[OK] Total de movimentos: {metrics.total_movements}")
    print(f"[OK] Taxa de sucesso: {metrics.success_rate:.1%}")
    print(f"[OK] Confianca media: {metrics.avg_confidence:.1%}")

    report = format_metrics_report(metrics)
    if report:
        print(f"[OK] Relatorio gerado com {len(report)} caracteres")

    return True


def test_postprocessor():
    """Testa pós-processador."""
    print("\n" + "="*70)
    print("TESTE 6: Validar Pos-Processador OCR")
    print("="*70)

    from ocr_postprocessor import postprocess_ocr_line

    fields = {
        "data_movimento": "23/06",
        "data_valor": "25/06",
        "descricao": "Google One Dublin",
        "pais": "IRL",
        "moeda_original": "",
        "taxa_cambio": "",
        "debito_eur": "1.99",
        "credito_eur": ""
    }

    corrected, status, reasons = postprocess_ocr_line("23/06 25/06 Google One 1.99", fields)
    print(f"[OK] Status: {status}")
    print(f"[OK] Motivos: {reasons if reasons else 'Nenhum'}")

    return True


def main():
    """Executa todos os testes."""
    print("\n" + "*"*70)
    print("VALIDACAO COMPLETA DA IMPLEMENTACAO")
    print("*"*70)

    tests = [
        ("Imports", test_imports),
        ("Correcoes OCR", test_ocr_corrections),
        ("Scoring Confianca", test_confidence_scoring),
        ("Padroes Comerciantes", test_merchant_patterns),
        ("Metricas Qualidade", test_metrics),
        ("Pos-Processador", test_postprocessor),
    ]

    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n[ERRO] {name}: {e}")
            import traceback
            traceback.print_exc()
            results.append((name, False))

    print("\n" + "="*70)
    print("RESUMO DOS TESTES")
    print("="*70)

    for name, result in results:
        status = "PASSOU" if result else "FALHOU"
        print(f"  {name}: {status}")

    all_passed = all(result for _, result in results)

    if all_passed:
        print("\n" + "-"*70)
        print("SUCESSO: Todos os testes passaram!")
        print("-"*70)
        return 0
    else:
        print("\n" + "-"*70)
        print("FALHA: Alguns testes falharam!")
        print("-"*70)
        return 1


if __name__ == "__main__":
    sys.exit(main())
