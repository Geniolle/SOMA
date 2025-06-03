###################################################################################
# Bloco: Verificação e download do ChromeDriver atualizado
###################################################################################

import os
import re
import zipfile
import requests
import shutil
import winreg
from io import BytesIO

def obter_versao_chrome():
    caminho = r"SOFTWARE\Google\Chrome\BLBeacon"
    try:
        reg = winreg.OpenKey(winreg.HKEY_CURRENT_USER, caminho)
        versao, _ = winreg.QueryValueEx(reg, "version")
        winreg.CloseKey(reg)
        return versao
    except Exception as e:
        print(f"Erro ao obter versão do Chrome: {e}")
        return None

def obter_url_chromedriver(versao_chrome):
    base = ".".join(versao_chrome.split(".")[:3])
    index_url = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build-with-downloads.json"
    try:
        resposta = requests.get(index_url)
        dados = resposta.json()
        patch = dados["builds"].get(base)
        if not patch:
            print(f"Nenhuma versão encontrada para base {base}")
            return None
        for item in patch["downloads"]["chromedriver"]:
            if item["platform"] == "win64":
                return item["url"]
        print("Versão para win64 não encontrada.")
        return None
    except Exception as e:
        print(f"Erro ao obter URL do ChromeDriver: {e}")
        return None

def download_e_extrair_chromedriver(url, destino):
    try:
        resposta = requests.get(url)
        with zipfile.ZipFile(BytesIO(resposta.content)) as zip_ref:
            for membro in zip_ref.namelist():
                if "chromedriver.exe" in membro:
                    zip_ref.extract(membro, destino)
                    origem = os.path.join(destino, membro)
                    destino_final = os.path.join(destino, "chromedriver.exe")
                    if os.path.exists(destino_final):
                        os.remove(destino_final)
                    shutil.move(origem, destino_final)
                    print(f"✅ ChromeDriver atualizado em: {destino_final}")
                    break
    except Exception as e:
        print(f"Erro ao extrair ChromeDriver: {e}")

def verificar_ou_atualizar_chromedriver(destino):
    versao = obter_versao_chrome()
    if not versao:
        print("❌ Não foi possível obter a versão do Chrome.")
        return
    print(f"🧭 Versão atual do Google Chrome: {versao}")
    chromedriver_path = os.path.join(destino, "chromedriver.exe")
    if os.path.exists(chromedriver_path):
        print(f"📂 ChromeDriver já existe em: {chromedriver_path}")
    else:
        print("🔽 ChromeDriver não encontrado. Iniciando download...")

    url = obter_url_chromedriver(versao)
    if url:
        print(f"🔗 URL de download: {url}")
        download_e_extrair_chromedriver(url, destino)
    else:
        print("❌ Não foi possível obter o link de download do ChromeDriver.")

# Executa verificação e download antes do WebDriver ser iniciado
pasta_chromedriver = r"G:\O meu disco\python\chromedriver"
verificar_ou_atualizar_chromedriver(pasta_chromedriver)
