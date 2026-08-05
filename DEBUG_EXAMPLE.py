#!/usr/bin/env python3
"""
Exemplo de como usar o sistema de debug interativo do SOMA.

Para testar:
1. Ativar: DEBUG_SELECTOR_INTERACTIVE=true no deploy/.env
2. Executar: python DEBUG_EXAMPLE.py
3. Pressionar ENTER em cada pausa para acompanhar
"""

import os
import sys
from pathlib import Path

# Setup paths
sys.path.insert(0, str(Path(__file__).parent / "src"))

from soma_app.infra.env import env_bool
from soma_app.infra.log_config import configure_logging
from soma_app.infra.webdriver_factory import create_bundle
from soma_app.config.settings import Settings
from selenium.webdriver.common.by import By
import logging

log = logging.getLogger(__name__)


def main():
    """Demonstra o sistema de debug interativo."""

    # Setup logging
    log_path, audit_path = configure_logging()
    log.info("=" * 70)
    log.info("DEMO: Sistema de Debug Interativo do SOMA")
    log.info("=" * 70)

    # Verificar se debug está ativo
    debug_active = env_bool("DEBUG_SELECTOR_INTERACTIVE", default=False)
    log.info(f"Mode DEBUG_SELECTOR_INTERACTIVE: {debug_active}")

    if not debug_active:
        log.warning("⚠️  Debug interativo NÃO está ativo!")
        log.warning("Para ativar, edite deploy/.env:")
        log.warning("  DEBUG_SELECTOR_INTERACTIVE=true")
        print("\n✗ Debug não está ativado. Saindo.")
        return

    log.info("✓ Debug interativo ATIVO")
    print("\n" + "=" * 70)
    print("✓ Sistema de Debug Interativo Ativado")
    print("=" * 70)

    # Criar bundle do WebDriver
    print("\n1️⃣  Criando WebDriver...")
    try:
        bundle = create_bundle()
        driver = bundle.driver
        actions = bundle.a

        if not actions:
            log.error("Falha ao criar Actions")
            return

        log.info("✓ WebDriver criado com sucesso")
        print("✓ WebDriver iniciado")

        # Exemplo 1: Screenshot
        print("\n2️⃣  Tirando screenshot de teste...")
        screenshot_path = actions.screenshot("debug_example_start")
        log.info(f"Screenshot salvo: {screenshot_path}")
        print(f"✓ Screenshot: {screenshot_path}")

        # Exemplo 2: Demonstrar logs
        print("\n3️⃣  Demonstrando logs de ação...")
        print("\nQuando você fizer uma ação (click, type, etc), verá:")
        print("  [SELECTOR] action=... | method=... | selector=...")
        print("  ✓ action=... | method=... | selector=...")
        print("  → Pressione ENTER para continuar...")

        # Exemplo 3: Teste um seletor
        print("\n4️⃣  Testando um seletor simples...")
        test_locator = (By.TAG_NAME, "body")

        print(f"\nBuscando elemento: {test_locator}")
        try:
            element = actions.wait_present(test_locator, timeout_seconds=5)
            log.info(f"✓ Elemento encontrado: {element.tag_name}")
            print(f"✓ Elemento encontrado: {element.tag_name}")
        except Exception as e:
            log.error(f"✗ Erro ao encontrar elemento: {e}")
            print(f"✗ Erro: {e}")

        # Exemplo 4: Timeout
        print("\n5️⃣  Simulando um timeout (seletor que não existe)...")
        fake_locator = (By.ID, "elemento_que_nao_existe_12345")
        print(f"\nBuscando elemento inexistente: {fake_locator}")
        try:
            element = actions.wait_present(fake_locator, timeout_seconds=2)
            print(f"✓ Elemento encontrado: {element.tag_name}")
        except Exception as e:
            print(f"✓ Timeout capturado (esperado): {type(e).__name__}")
            print(f"\nArquivos gerados:")
            print(f"  📷 Screenshot: artifacts/screenshots/timeout_*.png")
            print(f"  📄 HTML: artifacts/diagnostics/timeout_*.html")
            print(f"  📋 Log: logs/soma_selectors_*.log")

        # Resumo
        print("\n" + "=" * 70)
        print("📊 RESUMO DOS ARQUIVOS GERADOS")
        print("=" * 70)
        print(f"\n📂 Logs:")
        print(f"   {log_path}")

        log_dir = Path(log_path).parent
        selector_logs = list(log_dir.glob("soma_selectors_*.log"))
        if selector_logs:
            print(f"\n📋 Log de Seletores:")
            for log_file in selector_logs[-1:]:  # Última arquivo
                print(f"   {log_file}")
                print(f"\n   Conteúdo (últimas linhas):")
                with open(log_file) as f:
                    lines = f.readlines()
                    for line in lines[-5:]:
                        print(f"   {line.rstrip()}")

        print(f"\n📁 Artifacts:")
        artifacts = Path("artifacts")
        if artifacts.exists():
            screenshots = list(artifacts.glob("screenshots/*.png"))
            if screenshots:
                print(f"   📷 Screenshots: {len(screenshots)} arquivo(s)")
                for ss in screenshots[-3:]:
                    print(f"      - {ss.name}")

            diagnostics = list(artifacts.glob("diagnostics/*.html"))
            if diagnostics:
                print(f"   📄 Diagnostics: {len(diagnostics)} arquivo(s)")
                for diag in diagnostics[-3:]:
                    print(f"      - {diag.name}")

        print("\n" + "=" * 70)
        print("✓ Demo concluído com sucesso!")
        print("=" * 70)

        # Cleanup
        driver.quit()
        log.info("WebDriver finalizado")

    except Exception as e:
        log.exception(f"Erro na demo: {e}")
        print(f"\n✗ Erro: {e}")
        raise


if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("🐛 DEMO: Sistema de Debug Interativo do SOMA")
    print("=" * 70)
    print("\nEste script demonstra as capacidades de debug do sistema.")
    print("Você verá logs, screenshots e HTML diagnostics.")
    print("\nPressione ENTER em cada pausa ou Ctrl+C para sair.")
    print("=" * 70 + "\n")

    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️  Interrompido pelo usuário")
        sys.exit(0)
