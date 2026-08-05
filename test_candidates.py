#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Teste Automatico: Verificar qual candidato funciona
Simula o comportamento de wait_any_present sem precisar rodar a app
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage

print("\n" + "=" * 80)
print("TESTE AUTOMATICO DE CANDIDATOS: BTN_INSERIR_BAIXA_CANDIDATES")
print("=" * 80)

candidates = EntradasSaidasPage.BTN_INSERIR_BAIXA_CANDIDATES

print(f"\nEncontrados {len(candidates)} candidatos:")
print("-" * 80)

for i, (by_type, selector) in enumerate(candidates, 1):
    print(f"\n[{i}] Candidato {i}")
    print(f"    Tipo: {by_type}")
    print(f"    Seletor: {selector[:70]}...")

print("\n" + "-" * 80)
print("\nNOTA: Para testar os candidatos de verdade, e preciso do Selenium +")
print("      navegador aberto. O sistema wait_any_present() tentara cada um")
print("      ate encontrar que funciona.")
print("\nO sistema agora esta configurado para:")
print("  1. Tentar candidato 1")
print("  2. Se falhar, tentar candidato 2")
print("  3. Continuar ate encontrar que funciona")
print("  4. Usar o primeiro que funciona")
print("\n" + "=" * 80)
print("PROXIMOS PASSOS:")
print("  1. Executar: python main.py")
print("  2. Sistema vai tentar os 5 candidatos")
print("  3. Quando der certo, o primeiro candidato que funcionar sera usado")
print("=" * 80)
