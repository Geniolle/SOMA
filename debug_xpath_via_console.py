#!/usr/bin/env python3
"""
Script que executa SOMA, aguarda chegar ao modal, e depois usa um console script
para colectar os XPaths dos elementos críticos.
"""
import subprocess
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

# Iniciar SOMA em background com HEADLESS=false
print("🚀 Iniciando SOMA com Chrome visível...")
env = os.environ.copy()
env['HEADLESS'] = 'false'

soma_process = subprocess.Popen(
    ['python', 'main.py'],
    cwd='C:\\workspace\\SOMA',
    env=env,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE
)

print(f"✓ SOMA iniciado (PID: {soma_process.pid})")
print("⏳ Aguardando 40 segundos até chegar ao ponto crítico...")

# Aguardar que Chrome abra
time.sleep(40)

# Agora vamos tentar conectar e colectar os elementos
print("\n🔍 Conectando ao Chrome para análise...")

try:
    # Inicializar driver conectando ao Chrome existente
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)

    # Tentar encontrar a porta do Chrome debugger
    print("Procurando Chrome aberto...")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

    print(f"✓ Conectado! URL: {driver.current_url}")
    time.sleep(2)

    # Executar um script JavaScript para colectar os elementos
    js_code = """
    return {
        'page_title': document.title,
        'page_url': window.location.href,
        'modals': document.querySelectorAll('[role="dialog"], [role="alertdialog"], .modal, .popup, .swal').length,
        'buttons_ok': Array.from(document.querySelectorAll('button')).filter(b => b.textContent.includes('OK')).map(b => ({text: b.textContent, class: b.className})),
        'buttons_sim': Array.from(document.querySelectorAll('button')).filter(b => b.textContent.includes('Sim')).map(b => ({text: b.textContent, class: b.className})),
        'inputs': Array.from(document.querySelectorAll('input')).map(i => ({name: i.name, placeholder: i.placeholder, type: i.type}))
    }
    """

    result = driver.execute_script(js_code)
    print("\n📋 Análise do DOM:")
    import json
    print(json.dumps(result, indent=2, ensure_ascii=False))

    # Tirar screenshot
    print("\n📸 Capturando screenshot...")
    driver.save_screenshot('C:\\workspace\\SOMA\\artifacts\\screenshots\\debug_modal.png')
    print("✓ Screenshot salvo")

    driver.quit()

except Exception as e:
    print(f"❌ Erro ao conectar: {e}")
    import traceback
    traceback.print_exc()

# Finalizar SOMA
print("\n🛑 Finalizando SOMA...")
soma_process.terminate()
soma_process.wait(timeout=10)
print("✓ Concluído")
