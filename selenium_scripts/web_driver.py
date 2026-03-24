import os
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from config.config import HEADLESS

def iniciar_webdriver():
    """
    Inicializa o WebDriver com configurações otimizadas para evitar o modo mobile
    e garantir compatibilidade com servidores Linux.
    """
    chrome_options = Options()
    
    if HEADLESS:
        chrome_options.add_argument("--headless=new")
    
    # === A CURA PARA A CEGUEIRA NO SERVIDOR LINUX ===
    # Forçamos uma resolução Full HD e um User-Agent de Desktop
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--force-device-scale-factor=1")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # Otimizações para ambiente Linux/Server (essencial para evitar crashes)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Instalação automática do Driver compatível
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    # Maximizar a janela para garantir que todos os elementos estão visíveis
    driver.maximize_window()
    
    return driver

def salvar_screenshot(driver, nome_arquivo):
    """
    Função auxiliar utilizada pelo login.py e outros scripts para capturar 
    a tela em caso de erro ou para auditoria.
    """
    try:
        # Cria a pasta logs se ela não existir
        pasta_logs = Path("logs")
        pasta_logs.mkdir(exist_ok=True)
        
        caminho_final = pasta_logs / nome_arquivo
        driver.save_screenshot(str(caminho_final))
        print(f"📸 Screenshot salva em: {caminho_final}")
    except Exception as e:
        print(f"❌ Falha ao salvar screenshot: {e}")