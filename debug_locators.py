#!/usr/bin/env python3
"""
Script de Debug: Verificar se locators estão sendo carregados corretamente.

Para usar no VSCode:
1. Abrir terminal integrado (Ctrl+`)
2. Executar: python debug_locators.py
3. Adicionar breakpoint se necessário
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from soma_app.automation.actions import Actions, ActionConfig
from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.config.settings import Settings
from soma_app.config.locators import load_page_locator_config

print("=" * 80)
print("🔍 DEBUG: Verificar Locators BTN_INSERIR_BAIXA")
print("=" * 80)

# 1. Carregar configuração de locators do JSON
print("\n1️⃣ Carregando locators do JSON...")
cfg = load_page_locator_config(None, "entradas_saidas")
print(f"   Locators carregados do JSON: {len(cfg)} items")

if "BTN_INSERIR_BAIXA_CANDIDATES" in cfg:
    print(f"\n   ✅ BTN_INSERIR_BAIXA_CANDIDATES encontrado no JSON!")
    candidates = cfg.get("BTN_INSERIR_BAIXA_CANDIDATES", [])
    print(f"   📋 Quantidade de candidates: {len(candidates)}")
    for i, candidate in enumerate(candidates, 1):
        print(f"      {i}. {candidate}")
else:
    print(f"\n   ❌ BTN_INSERIR_BAIXA_CANDIDATES NÃO encontrado no JSON!")
    print(f"   Chaves disponíveis: {list(cfg.keys())}")

# 2. Verificar classe EntradasSaidasPage antes de apply_locator_overrides
print("\n2️⃣ Verificando EntradasSaidasPage ANTES de apply_locator_overrides...")
print(f"   BTN_INSERIR_BAIXA_CANDIDATES (antes): {EntradasSaidasPage.BTN_INSERIR_BAIXA_CANDIDATES}")

# 3. Criar instância e aplicar locator overrides
print("\n3️⃣ Criando instância e aplicando locator overrides...")
try:
    # Criar Settings vazio
    settings = Settings()

    # Criar página (vai aplicar overrides no __init__)
    from unittest.mock import MagicMock
    mock_actions = MagicMock(spec=Actions)

    page = EntradasSaidasPage(mock_actions, settings)

    print(f"   ✅ Instância criada com sucesso!")
    print(f"   BTN_INSERIR_BAIXA_CANDIDATES (depois): {page.BTN_INSERIR_BAIXA_CANDIDATES}")
    print(f"   Quantidade: {len(page.BTN_INSERIR_BAIXA_CANDIDATES)}")

    if page.BTN_INSERIR_BAIXA_CANDIDATES:
        print(f"\n   ✅ CANDIDATES FORAM CARREGADOS!")
        for i, candidate in enumerate(page.BTN_INSERIR_BAIXA_CANDIDATES, 1):
            print(f"      {i}. {candidate}")
    else:
        print(f"\n   ❌ CANDIDATES ESTÃO VAZIOS!")
        print(f"   Verifique o JSON em locators.json")

except Exception as e:
    print(f"   ❌ Erro ao criar instância: {e}")
    import traceback
    traceback.print_exc()

# 4. Informações úteis para debug
print("\n4️⃣ Informações úteis para debug:")
print(f"   • Arquivo de locators: {Path('src/soma_app/config/locators.json')}")
print(f"   • Classe: {Path('src/soma_app/automation/pages/entradas_saidas_page.py')}")
print(f"   • Arquivo de config: {Path('src/soma_app/config/locators.py')}")

print("\n" + "=" * 80)
print("✅ Debug concluído!")
print("=" * 80)
print("\n💡 Dicas para usar no VSCode:")
print("   1. Abra Debug (Ctrl+Shift+D)")
print("   2. Selecione 'SOMA Debug'")
print("   3. Clique em Play ou pressione F5")
print("   4. Adicione breakpoints clicando na linha")
print("   5. Use Debug Console (Ctrl+Shift+Y) para inspecionar variáveis")
