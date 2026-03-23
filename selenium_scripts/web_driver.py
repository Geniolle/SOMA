from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from config.config import HEADLESS

def iniciar_webdriver():
    chrome_options = Options()
    
    if HEADLESS:
        chrome_options.add_argument("--headless=new") # O novo modo headless (mais estável)
    
    # === A CURA PARA A CEGUEIRA NO SERVIDOR LINUX ===
    chrome_options.add_argument("--window-size=1920,1080") # Forçar monitor Full HD
    chrome_options.add_argument("--force-device-scale-factor=1")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # Otimizações para Linux
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # Maximizar a janela só para garantir
    driver.maximize_window()
    
    return driver