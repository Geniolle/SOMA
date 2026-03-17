###################################################################################
# CONFIGURAÇÕES GERAIS DO PROJETO SOMA
###################################################################################

import os

# URL do site do SOMA
SOMA_URL_INICIAL = "https://verbodavida.info/apps/index.php"

# Credenciais do SOMA
EMAIL_SOMA = "familialopesemportugal@gmail.com"
SENHA_SOMA = "P@1internet"

# Timeout padrão do Selenium
SELENIUM_TIMEOUT = 20

# Diretório de logs
LOGS_DIR = os.path.join(os.getcwd(), "logs")
os.makedirs(LOGS_DIR, exist_ok=True)

# Caminho para o ficheiro de credenciais do Google Sheets
CREDENCIAIS_PATH = os.path.join(os.getcwd(), "config", "credenciais.json")

# URL da planilha do Google Sheets
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing"

# Nome da worksheet
SHEET_NAME = "CONTAORDEM"
