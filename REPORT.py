import gspread
import sys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import locale
import time
import os
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Configurar o logger
logging.basicConfig(level=logging.INFO, format='%(message)s')
def log_info(message):
    logging.info(message)

# Configurar o locale para português
locale.setlocale(locale.LC_TIME, 'pt_PT.UTF-8')

# (1)#################################################################################
print()
print("==================================================================")
print(f"(1) Iniciando o processo")
print("==================================================================")

# Configuração para acesso ao Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciais_path = os.path.join(os.path.dirname(__file__), 'credenciais.json')

# Autenticação
credentials = ServiceAccountCredentials.from_json_keyfile_name(credenciais_path, scope)
client = gspread.authorize(credentials)

# Abrir a planilha e selecionar as sheets
sheet_controller_name = "CONTROLLER"
sheet_email_name = "EMAIL"

def carregar_sheet(sheet_name):
    return client.open_by_url("https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing").worksheet(sheet_name)

sheet_controller = carregar_sheet(sheet_controller_name)
sheet_email = carregar_sheet(sheet_email_name)

# Configuração do WebDriver para o Chrome sem modo headless para visualização
chrome_options = Options()

# Remover o modo headless para execução visual
# chrome_options.add_argument("--headless")  # Comente esta linha para habilitar a visualização
chrome_options.add_argument("--disable-gpu")  # Otimizar para headless
chrome_options.add_argument("--no-sandbox")  # Necessário em alguns sistemas
chrome_options.add_argument("--disable-dev-shm-usage")  # Melhor para memória compartilhada
chrome_options.add_argument("--window-size=1920,1080")  # Tamanho da janela para evitar problemas de layout

webdriver_path = "G:\\O meu disco\\python\\chromedriver\\chromedriver.exe"  # Caminho atualizado para o ChromeDriver
service = Service(webdriver_path)

# Inicializar o WebDriver no modo visual (sem headless)
try:
    driver = webdriver.Chrome(service=service, options=chrome_options)
    print("WebDriver inicializado com sucesso em modo visual.")
except Exception as e:
    print(f"Erro ao iniciar o WebDriver: {e}")
    exit(1)  # Finalizar o script em caso de falha

"""
# Configuração do WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
webdriver_path = "G:\\O meu disco\\python\\chromedriver\\chromedriver.exe"
driver = webdriver.Chrome(service=Service(webdriver_path), options=chrome_options)
"""

# (2)#################################################################################
print()
print("==================================================================")
print(f"(2) Iniciando o processo de acesso ao SOMA")
print("==================================================================")

url = 'https://verbodavida.info/apps/index.php'
driver.get(url)

def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
    except:
        print(f"Elemento não encontrado ou visível: {valor}")
        driver.save_screenshot("erro.png")
        return None

usuario = tentar_encontrar_elemento(driver, By.NAME, 'email')
if usuario:
    usuario.send_keys('familialopesemportugal@gmail.com')
    print(f"Preenchimento do utilizador realizado com sucesso")

senha = tentar_encontrar_elemento(driver, By.NAME, 'senha')
if senha:
    senha.send_keys('P@1internet')
    print("Preenchimento da senha realizado com sucesso")

botao_login = tentar_encontrar_elemento(driver, By.NAME, 'submit')
if botao_login:
    botao_login.click()
    print("Click do botão login realizado com sucesso")

try:
    botao_soma = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, '285'))
    )
    botao_soma.click()
    print("Botão 'SOMA' clicado com sucesso!")
except TimeoutException:
    print("Erro: Botão 'SOMA' não encontrado ou não ficou disponível a tempo!")

try:
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
    )
    print("Página carregada com sucesso após clicar no botão 'SOMA'.")
except TimeoutException:
    print("Erro: Página não carregou após clicar no botão 'SOMA'.")

# (3)#################################################################################
print()
print("==================================================================")
print(f"(3) Iniciando o processamento de dados da sheet {sheet_controller_name}")
print("==================================================================")

# Carregar os dados da planilha
controller_data = sheet_controller.get_all_records()
print(f"Dados carregados com sucesso da sheet {sheet_controller_name}.")

# Filtrar as linhas
linhas_filtradas = [
    (index, row)
    for index, row in enumerate(controller_data, start=2)
    if row.get('FECHAMENTO', '').strip().upper() == 'CONCLUÍDO' and str(row.get('STATUS PROCESSAMENTO', '')).strip() in ['', 'Em processamento']
]

print(f"Linhas filtradas encontradas na sheet {sheet_controller_name}: {len(linhas_filtradas)}")

# Processar as linhas filtradas
for index, row in linhas_filtradas:
    try:
        print(f"Processando linha {index}: {row}")
        periodo = row.get('PERÍODO', '').strip().capitalize()
        ano_fiscal = row.get('ANO FISCAL', '')

        # Criar intervalo de datas
        data_inicial = datetime.strptime(f"1 {periodo} {ano_fiscal}", "%d %B %Y")
        ultimo_dia = (data_inicial.replace(month=data_inicial.month % 12 + 1, day=1) - timedelta(days=1)).day
        data_final = data_inicial.replace(day=ultimo_dia)
        print(f"Intervalo de datas para o PERÍODO '{periodo}' e ANO FISCAL '{ano_fiscal}':")
        print(f"- Data inicial: {data_inicial.strftime('%d/%m/%Y')}")
        print(f"- Data final: {data_final.strftime('%d/%m/%Y')}")

        # Gerar URL do relatório
        relatorio_url = (
            f"https://verbodavida.info/IVV/sys/relatorios/relatorios_fluxoCaixa.php"
            f"?t=2&ft=1,5,2,7,3,4,6&i=270&c={data_inicial.strftime('%d/%m/%Y')}"
            f"&f={data_final.strftime('%d/%m/%Y')}&id_c=null&t_d=1&id_cc=-1&id_m=2&o=undefined"
        )
        print(f"URL do relatório: {relatorio_url}")

        # Configurar o Chrome para salvar automaticamente o PDF
        download_dir = os.getcwd()  # Diretório atual
        relatorio_pdf_path = os.path.join(download_dir, f"relatorio_{ano_fiscal}_{str(data_inicial.month).zfill(2)}.pdf")
        
        # Atualizar opções do Chrome para salvar PDF automaticamente
        chrome_prefs = {
            "printing.print_preview_sticky_settings.appState": json.dumps({
                "version": 2,
                "isHeaderFooterEnabled": False,
                "selectedDestinationId": "Save as PDF",
                "mediaSize": {},
                "marginsType": 1,
                "isCssBackgroundEnabled": True
            }),
            "savefile.default_directory": download_dir
        }
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": download_dir})
        driver.execute_script('window.print();')  # Executar a ação de imprimir como PDF

        # Verificar se o arquivo foi salvo
        time.sleep(5)  # Aguardar o tempo de salvar
        if os.path.exists(relatorio_pdf_path):
            print(f"Relatório salvo em PDF com sucesso: {relatorio_pdf_path}")
        else:
            print(f"Erro: Relatório PDF não foi encontrado no caminho esperado: {relatorio_pdf_path}")
            raise FileNotFoundError(f"Relatório não encontrado no caminho: {relatorio_pdf_path}")

    except Exception as e:
        print(f"Erro inesperado ao processar a linha {index}: {e}")






# (4)#################################################################################
print()
print("==================================================================")
print(f"(4) Processando dados da sheet {sheet_email_name}")
print("==================================================================")

email_data = sheet_email.get_all_records()
linhas_filtradas_email = [
    (index, row)
    for index, row in enumerate(email_data, start=2)
    if row.get('EMAIL', '').strip() != '' and str(row.get('STATUS PROCESSAMENTO', '')).strip() in ['', 'Em processamento']
]

print(f"Linhas filtradas encontradas na sheet {sheet_email_name}: {len(linhas_filtradas_email)}")

for index, row in linhas_filtradas_email:
    try:
        email = row.get('EMAIL', '').strip()
        print(f"Enviando e-mail para: {email}")

        # Configuração do servidor SMTP
        smtp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = "verbodavidabraga@gmail.com"
        sender_password = "P@1geniole01253121"

        # Configurar o e-mail
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = email
        message['Subject'] = "Relatório Mensal"
        message.attach(MIMEText("Segue em anexo o relatório mensal.", 'plain'))

        # Anexar PDF
        pdf_path = os.path.join(os.getcwd(), f"relatorio_{ano_fiscal}_{str(data_inicial.month).zfill(2)}.pdf")  # Caminho atualizado para o PDF
        try:
            with open(pdf_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename={os.path.basename(pdf_path)}",
                )
                message.attach(part)
        except FileNotFoundError:
            print(f"Erro: Arquivo PDF não encontrado: {pdf_path}")
            try:
                col_status = list(email_data[0].keys()).index('STATUS PROCESSAMENTO') + 1
                sheet_email.update_cell(index, col_status, "Erro: PDF não encontrado")
            except Exception as e:
                print(f"Erro ao atualizar a planilha: {e}")
            continue

        # Enviar o e-mail
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, email, message.as_string())
                print(f"E-mail enviado com sucesso para: {email}")
                try:
                    col_status = list(email_data[0].keys()).index('STATUS PROCESSAMENTO') + 1
                    sheet_email.update_cell(index, col_status, "Enviado")
                except Exception as e:
                    print(f"Erro ao atualizar a planilha: {e}")
        except Exception as e:
            print(f"Erro ao enviar e-mail para: {email}. Detalhes: {e}")
            try:
                col_status = list(email_data[0].keys()).index('STATUS PROCESSAMENTO') + 1
                sheet_email.update_cell(index, col_status, "Erro ao enviar e-mail")
            except Exception as e:
                print(f"Erro ao atualizar a planilha: {e}")

    except Exception as e:
        print(f"Erro ao processar a linha {index}: {e}")
        try:
            col_status = list(email_data[0].keys()).index('STATUS PROCESSAMENTO') + 1
            sheet_email.update_cell(index, col_status, "Erro")
        except Exception as e:
            print(f"Erro ao atualizar a planilha: {e}")


# (5)#################################################################################
print()
print("==================================================================")
print("Finalizando o processo")
print("==================================================================")

# Fechar o WebDriver
try:
    driver.quit()
    print("WebDriver finalizado com sucesso.")
except Exception as e:
    print(f"Erro ao finalizar o WebDriver: {e}")

print("Processo concluído com sucesso!")
sys.exit(0)
