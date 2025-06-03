import os

estrutura = [
    "config",
    "sheets",
    "selenium_scripts",
    "logs"
]

ficheiros = [
    "config/config.py",
    "sheets/sheets_service.py",
    "sheets/atualizar_caixas.py",
    "selenium_scripts/web_driver.py",
    "selenium_scripts/login.py",
    "selenium_scripts/navegacao.py",
    "selenium_scripts/processar_saida.py",
    "selenium_scripts/processar_entrada.py",
    "selenium_scripts/processar_transferencia.py",
    "selenium_scripts/utilitarios.py",
    "main.py",
    "requirements.txt"
]

for pasta in estrutura:
    os.makedirs(pasta, exist_ok=True)

for ficheiro in ficheiros:
    open(ficheiro, "a").close()

print("Estrutura criada com sucesso!")
