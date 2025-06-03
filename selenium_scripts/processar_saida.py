###################################################################################
# selenium_scripts/processar_saida.py
###################################################################################

from selenium.webdriver.common.by import By
import time
from selenium_scripts.navegacao import (
    tentar_encontrar_elemento,
    clicar_no_elemento,
    redirecionar_para_inicio
)
from selenium_scripts.web_driver import salvar_screenshot

def aceder_pagina_entradas_saidas(driver):
    print("📄 A aceder à página 'Entradas/saídas'...")
    try:
        redirecionar_para_inicio(driver)
        time.sleep(3)
        elemento = driver.find_element(By.XPATH, "//span[text()='Entradas/saídas']")
        elemento.click()
        print("✅ Página 'Entradas/saídas' carregada com sucesso!")
        time.sleep(3)
        return True
    except:
        print("❌ Não foi possível encontrar o botão 'Entradas/saídas'.")
        salvar_screenshot(driver, "erro_elemento.png")
        return False

def processar_saida(driver, linha, idx, sheet):
    print(f"\n🟡 Iniciando o processamento de SAÍDA")
    print(f"🔎 Linha: {idx}, Dados: {linha}")

    if not aceder_pagina_entradas_saidas(driver):
        return

    try:
        print("📝 Preenchendo campos da saída...")

        # PLANO DE CONTA
        campo_plano = tentar_encontrar_elemento(driver, By.XPATH, "//label[text()='Plano de Conta']/following-sibling::span//input")
        if campo_plano:
            campo_plano.send_keys(linha['PLANO DE CONTA'])
            time.sleep(1)

        # CENTRO DE CUSTO
        campo_cc = tentar_encontrar_elemento(driver, By.XPATH, "//label[text()='Centro de Custo']/following-sibling::span//input")
        if campo_cc:
            campo_cc.send_keys(linha['CENTRO DE CUSTO'])
            time.sleep(1)

        # DESCRIÇÃO SOMA
        campo_desc = tentar_encontrar_elemento(driver, By.XPATH, "//label[text()='Descrição']/following-sibling::input")
        if campo_desc:
            campo_desc.clear()
            campo_desc.send_keys(linha['DESCRIÇÃO SOMA'])
            time.sleep(1)

        # IMPORTÂNCIA
        campo_valor = tentar_encontrar_elemento(driver, By.XPATH, "//label[text()='Importância']/following-sibling::input")
        if campo_valor:
            campo_valor.clear()
            campo_valor.send_keys(str(linha['IMPORTÂNCIA']))
            time.sleep(1)

        # FORMA DE PAGAMENTO
        campo_pag = tentar_encontrar_elemento(driver, By.XPATH, "//label[text()='Forma de pagamento']/following-sibling::span//input")
        if campo_pag:
            campo_pag.send_keys(linha['FORMA DE PAGAMENTO'])
            time.sleep(1)

        # CONTA DE SAÍDA
        campo_caixa = tentar_encontrar_elemento(driver, By.XPATH, "//label[text()='Conta de Saída']/following-sibling::span//input")
        if campo_caixa:
            campo_caixa.send_keys(linha['CAIXA'])
            time.sleep(1)

        print("✅ Todos os campos foram preenchidos com sucesso!")

    except Exception as e:
        print(f"❌ Erro durante o preenchimento dos campos: {e}")
        salvar_screenshot(driver, "erro_processamento_saida.png")
