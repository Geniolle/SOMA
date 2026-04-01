###################################################################################
# FUNÇÃO DE LOGIN NO SISTEMA SOMA (VERSÃO SINCRONIZADA)
###################################################################################

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time

from config.config import SOMA_URL_INICIAL, EMAIL_SOMA, SENHA_SOMA, SELENIUM_TIMEOUT
from selenium_scripts.web_driver import salvar_screenshot


def tentar_encontrar_elemento(driver, by, valor, timeout=10):
    """Função auxiliar local para evitar erros de importação circular"""
    try:
        return WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((by, valor))
        )
    except Exception:
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
            raise TimeoutException("Campo 'email' não encontrado!")

        # Preencher senha
        senha_field = tentar_encontrar_elemento(driver, By.NAME, "senha", SELENIUM_TIMEOUT)
        if senha_field:
            senha_field.clear()
            senha_field.send_keys(SENHA_SOMA)
            print("   ✅ Senha preenchida.")
        else:
            print("   ❌ Campo 'senha' não encontrado!")
            raise TimeoutException("Campo 'senha' não encontrado!")

        # Clicar no botão Login
        login_button = tentar_encontrar_elemento(driver, By.NAME, "submit", SELENIUM_TIMEOUT)
        if login_button:
            login_button.click()
            print("   ✅ Login submetido, a aguardar o botão 'SOMA'.")
        else:
            print("   ❌ Botão 'Login' não encontrado!")
            raise TimeoutException("Botão 'Login' não encontrado!")

        time.sleep(2)

        # Clicar no botão SOMA na Central de Sistemas
        print("   ⏳ Aguardando botão 'SOMA' ficar clicável...")
        botao_soma = WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/main/div/div[2]/div/a"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_soma)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", botao_soma)
        print("   ✅ Botão 'SOMA' clicado com sucesso!")

        # Confirmar carregamento da área IVV
        WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            lambda d: "/ivv/" in d.current_url.lower()
        )
        print("   ✅ Área SOMA/IVV carregada com sucesso!")
        print(f"   URL atual: {driver.current_url}")

        time.sleep(2)

        # Clicar no atalho Entradas/saídas pelo onclick
        print("   ⏳ Aguardando atalho 'Entradas/saídas' ficar clicável...")
        atalho_entradas_saidas = WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            EC.element_to_be_clickable((
                By.XPATH,
                '//div[contains(@onclick, "exec=entradas_saidas") and .//span[contains(normalize-space(), "Entradas/saídas")]]'
            ))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", atalho_entradas_saidas)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", atalho_entradas_saidas)
        print("   ✅ Atalho 'Entradas/saídas' clicado com sucesso!")

        # Confirmação final da página de entradas/saídas
        WebDriverWait(driver, SELENIUM_TIMEOUT).until(
            lambda d: "exec=entradas_saidas" in d.current_url.lower()
        )
        print("✅ Página 'Entradas/saídas' carregada com sucesso!")
        print(f"   URL final: {driver.current_url}")

    except TimeoutException:
        print("❌ Erro no login: Timeout ao procurar elementos do SOMA.")
        print(f"   URL atual: {driver.current_url}")
        salvar_screenshot(driver, "erro_login_soma.png")
        raise
    except Exception as e:
        print(f"❌ Erro inesperado no login: {e}")
        print(f"   URL atual: {driver.current_url}")
        salvar_screenshot(driver, "erro_login_soma.png")
        raise