###################################################################################
# FUNÇÃO DE LOGIN NO SISTEMA SOMA (AJUSTADA)
###################################################################################

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
from selenium_scripts.web_driver import salvar_screenshot
from config.config import SOMA_URL_INICIAL, EMAIL_SOMA, SENHA_SOMA, SELENIUM_TIMEOUT
from selenium_scripts.navegacao import tentar_encontrar_elemento

def login_soma(driver):
    print("🔑 Iniciando login no sistema SOMA...")
    driver.get(SOMA_URL_INICIAL)

    try:
        # Preencher email
        email_field = tentar_encontrar_elemento(driver, By.NAME, "email")
        if email_field:
            email_field.send_keys(EMAIL_SOMA)
            print("✅ Email preenchido.")

        # Preencher senha
        senha_field = tentar_encontrar_elemento(driver, By.NAME, "senha")
        if senha_field:
            senha_field.send_keys(SENHA_SOMA)
            print("✅ Senha preenchida.")

        # Clicar no botão Login
        login_button = tentar_encontrar_elemento(driver, By.NAME, "submit")
        if login_button:
            login_button.click()
            print("✅ Login submetido, a aguardar o botão 'SOMA'.")

        # Clicar no botão 'SOMA' (ID 285)
        botao_soma = WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            EC.element_to_be_clickable((By.ID, "285"))
        )
        botao_soma.click()
        print("✅ Botão 'SOMA' clicado com sucesso!")

        # Confirmar carregamento da página "Entradas/Saídas"
        WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
        )
        print("✅ Página 'Entradas/saídas' carregada com sucesso!")

    except TimeoutException:
        print("❌ Erro no login: Timeout ao procurar elementos do SOMA.")
        salvar_screenshot(driver, "erro_login_soma.png")
        raise
