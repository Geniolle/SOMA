import sys
sys.path.insert(0, 'C:\\workspace\\SOMA\\src')

import time
import json
import subprocess
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options

print("=" * 80)
print("AUTO-CAPTURE XPATH - Coletando elementos críticos do SOMA")
print("=" * 80)

print("\n[1/5] Procurando Chrome em execução...")
try:
    result = subprocess.run(['tasklist', '/v'], capture_output=True, text=True, timeout=5)
    chrome_lines = [l for l in result.stdout.split('\n') if 'chrome' in l.lower()]
    print(f"  ✓ Encontrados {len(chrome_lines)} processo(s) Chrome")
except Exception as e:
    print(f"  ⚠ Erro ao procurar Chrome: {e}")

print("\n[2/5] Tentando conectar ao Chrome via debuggerAddress...")
driver = None
for port in [9222, 9223, 9224, 9225]:
    try:
        options = Options()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
        driver = webdriver.Chrome(options=options)
        print(f"  ✓ Conectado com sucesso na porta {port}!")
        break
    except Exception as e:
        pass

if not driver:
    print("\n✗ ERRO: Não consegui conectar ao Chrome")
    exit(1)

try:
    print("\n[3/5] Clicando em 'Sim' no popup...")
    time.sleep(2)
    
    try:
        sim_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Sim')]")
        sim_button.click()
        print("  ✓ Clicado em 'Sim'")
    except:
        print("  ⚠ Botão 'Sim' não encontrado")
    
    print("\n[4/5] Aguardando modal de pagamento...")
    time.sleep(4)
    
    print("\n[5/5] Capturando screenshot do modal...")
    screenshot_path = 'C:\\workspace\\SOMA\\artifacts\\screenshots\\modal_pagamento_capturado.png'
    driver.save_screenshot(screenshot_path)
    print(f"  ✓ Screenshot: {screenshot_path}")
    
    print("\n[6/5] Analisando página...")
    # Listar todos os inputs, buttons e modais
    page_source = driver.page_source
    
    # Procurar padrões de interesse
    if 'Adicione um pagamento' in page_source:
        print("  ✓ Encontrado: Modal 'Adicione um pagamento para esta conta'")
    else:
        print("  ✗ Modal 'Adicione um pagamento' não encontrado na página")
    
    if 'num_documento' in page_source or 'Documento' in page_source:
        print("  ✓ Encontrado: Campo 'Nº Documento' ou similar")
    else:
        print("  ✗ Campo 'Nº Documento' não encontrado")
    
    if 'Salvar' in page_source and 'Pagamento' in page_source:
        print("  ✓ Encontrado: Botão 'Salvar Pagamento' ou similar")
    else:
        print("  ✗ Botão 'Salvar Pagamento' não encontrado")
    
    if 'Inserir' in page_source and 'Baixa' in page_source:
        print("  ✓ Encontrado: Referência a 'Inserir Baixa'")
    else:
        print("  ✗ Referência a 'Inserir Baixa' não encontrada")
    
    print("\n" + "=" * 80)
    print("✓ Análise completa!")
    print("=" * 80)

finally:
    if driver:
        try:
            driver.quit()
        except:
            pass
