###################################################################################
# main.py
###################################################################################

from selenium_scripts.web_driver import iniciar_webdriver, salvar_screenshot
from selenium_scripts.login import login_soma
from selenium_scripts.processar_saida import processar_saida
from sheets.sheets_service import obter_sheet, obter_todos_os_registros
import time

print("🧪 Teste inicial do WebDriver, Google Sheets e processo SAÍDA")

# Inicializar WebDriver
driver = iniciar_webdriver()
print("✅ WebDriver inicializado com sucesso!")

# Login no sistema SOMA
login_soma(driver)

# Conectar à folha Google Sheets
sheet = obter_sheet()

# Obter todos os dados
registros = obter_todos_os_registros(sheet)

# Iterar pelas linhas
for idx, linha in enumerate(registros, start=2):
    if linha.get("TIPO", "").strip().upper() == "SAÍDA" and not linha.get("STATUS"):
        print(f"🔍 Encontrada linha para processamento: {linha}")
        try:
            processar_saida(driver, linha, idx, sheet)
        except Exception as e:
            print(f"❌ Erro ao processar a linha {idx}: {e}")
            salvar_screenshot(driver, "erro_geral.png")

# Encerrar WebDriver
driver.quit()
print("✅ WebDriver finalizado com sucesso.")
