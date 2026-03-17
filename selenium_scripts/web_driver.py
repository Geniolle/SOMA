###################################################################################
# SETUP DO SELENIUM WEBDRIVER
###################################################################################

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import os

from config.config import SELENIUM_TIMEOUT, LOGS_DIR

def iniciar_webdriver(headless=True):
    """
    Inicia e retorna o WebDriver do Chrome com as configurações padrão.
    """
    chrome_options = Options()

    if headless:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

    chrome_options.add_argument("--window-size=1920,1080")

    # Ocultar mensagens do DevTools no terminal
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

    try:
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        print("✅ WebDriver inicializado com sucesso!")
        return driver
    except Exception as e:
        print(f"❌ Erro ao iniciar o WebDriver: {e}")
        raise

def salvar_screenshot(driver, nome_arquivo="erro.png"):
    """
    Salva um screenshot do estado atual da página no diretório de logs.
    """
    os.makedirs(LOGS_DIR, exist_ok=True)
    caminho = os.path.join(LOGS_DIR, nome_arquivo)
    driver.save_screenshot(caminho)
    print(f"📸 Screenshot salva em: {caminho}")
