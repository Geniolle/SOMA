from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from config.config import HEADLESS

def iniciar_webdriver():
    chrome_options = Options()

    if HEADLESS:
        print("🤖 MODO: INVISÍVEL (Headless=True)")
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
    else:
        print("👁️ MODO: VISÍVEL (Headless=False)")
        # Quando False, a janela do Chrome abrirá fisicamente

    chrome_options.add_argument("--window-size=1920,1080")
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
    """Função mantida apenas para compatibilidade de importação."""
    pass