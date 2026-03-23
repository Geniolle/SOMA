###################################################################################
# FUNÇÃO DE LOGIN NO SISTEMA SOMA (VERSÃO SINCRONIZADA)
###################################################################################

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

# Importamos o que precisamos do nosso config central
from config.config import SOMA_URL_INICIAL, EMAIL_SOMA, SENHA_SOMA, SELENIUM_TIMEOUT
# Importamos a função de suporte do web_driver
from selenium_scripts.web_driver import salvar_screenshot

def tentar_encontrar_elemento(driver, by, valor, timeout=10):
    """Função auxiliar local para evitar erros de importação circular"""
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
    except:
        return None

def login_soma(driver):
    print("🔑 Iniciando login no sistema SOMA...")
    driver.get(SOMA_URL_INICIAL)

    try:
        # Preencher email
        email_field = tentar_encontrar_elemento(driver, By.NAME, "email", SELENIUM_TIMEOUT)
        if email_field:
            email_field.clear()
            email_field.send_keys(EMAIL_SOMA)
            print("   ✅ Email preenchido.")
        else:
            print("   ❌ Campo 'email' não encontrado!")

        # Preencher senha
        senha_field = tentar_encontrar_elemento(driver, By.NAME, "senha", SELENIUM_TIMEOUT)
        if senha_field:
            senha_field.clear()
            senha_field.send_keys(SENHA_SOMA)
            print("   ✅ Senha preenchida.")
        else:
            print("   ❌ Campo 'senha' não encontrado!")

        # Clicar no botão Login
        login_button = tentar_encontrar_elemento(driver, By.NAME, "submit", SELENIUM_TIMEOUT)
        if login_button:
            login_button.click()
            print("   ✅ Login submetido, a aguardar o botão 'SOMA'.")
        else:
            print("   ❌ Botão 'Login' não encontrado!")

        # Clicar no botão 'SOMA' (ID 285)
        print("   ⏳ Aguardando botão 'SOMA' ficar clicável...")
        botao_soma = WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            EC.element_to_be_clickable((By.ID, "285"))
        )
        botao_soma.click()
        print("   ✅ Botão 'SOMA' clicado com sucesso!")

        # Confirmar carregamento da página "Entradas/Saídas"
        WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
        )
        print("✅ Página 'Entradas/saídas' carregada com sucesso!")

    except TimeoutException:
        print("❌ Erro no login: Timeout ao procurar elementos do SOMA.")
        salvar_screenshot(driver, "erro_login_soma.png")
        raise
    except Exception as e:
        print(f"❌ Erro inesperado no login: {e}")
        raise