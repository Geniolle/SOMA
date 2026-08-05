#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto Debug: Testar candidatos automaticamente e corrigir
"""

import sys
import os
import json
from pathlib import Path

# Setup paths
sys.path.insert(0, str(Path(__file__).parent / "src"))
os.chdir(Path(__file__).parent)

from selenium.webdriver.common.by import By

print("\n" + "=" * 80)
print("AUTO DEBUG: Teste de Candidatos BTN_INSERIR_BAIXA")
print("=" * 80)

# 1. Carregar locators.json
print("\n[1] Carregando locators.json...")
locators_path = Path("src/soma_app/config/locators.json")
with open(locators_path, "r", encoding="utf-8") as f:
    data = json.load(f)

entradas_saidas_cfg = data.get("entradas_saidas", {})
candidates_json = entradas_saidas_cfg.get("BTN_INSERIR_BAIXA_CANDIDATES", [])

print(f"    Encontrados {len(candidates_json)} candidatos no JSON:")
for i, cand in enumerate(candidates_json, 1):
    print(f"    {i}. {cand[:60]}...")

# 2. Verificar classe EntradasSaidasPage
print("\n[2] Verificando classe EntradasSaidasPage...")
from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage

print(f"    BTN_INSERIR_BAIXA_CANDIDATES (classe): {EntradasSaidasPage.BTN_INSERIR_BAIXA_CANDIDATES}")

# 3. Carregar config de locators
print("\n[3] Testando apply_locator_overrides...")
from soma_app.config.locators import load_page_locator_config, apply_locator_overrides
from soma_app.config.settings import Settings
from unittest.mock import MagicMock
from soma_app.automation.actions import Actions

try:
    settings = Settings()
    mock_actions = MagicMock(spec=Actions)
    page = EntradasSaidasPage(mock_actions, settings)

    print(f"    BTN_INSERIR_BAIXA_CANDIDATES (depois): {page.BTN_INSERIR_BAIXA_CANDIDATES}")
    print(f"    Tamanho: {len(page.BTN_INSERIR_BAIXA_CANDIDATES)}")

    if page.BTN_INSERIR_BAIXA_CANDIDATES:
        print("\n    [OK] Candidatos foram carregados!")
        for i, cand in enumerate(page.BTN_INSERIR_BAIXA_CANDIDATES, 1):
            print(f"    {i}. {cand}")
    else:
        print("\n    [ERRO] Candidatos estao vazios!")
        print("    Problema: apply_locator_overrides nao preencheu a lista")

        # Tentar carregar manualmente
        print("\n[4] Tentando carregar manualmente do JSON...")
        cfg = load_page_locator_config(settings, "entradas_saidas")

        if "BTN_INSERIR_BAIXA_CANDIDATES" in cfg:
            print("    [OK] Encontrado no JSON!")
            items = cfg["BTN_INSERIR_BAIXA_CANDIDATES"]
            print(f"    {len(items)} items")
            for i, item in enumerate(items, 1):
                print(f"    {i}. {item}")
        else:
            print("    [ERRO] Nao encontrado no JSON!")
            print(f"    Chaves disponiveis: {list(cfg.keys())}")

except Exception as e:
    print(f"    [ERRO] Exception: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
print("AUTO DEBUG CONCLUIDO")
print("=" * 80)
