import gspread
from SYS_ENTRADA import atualizar_entradas
from SYS_SAÍDA import atualizar_saidas
from SYS_TRANSFERENCIA import atualizar_transferencias
from SYS_CAIXA import atualizar_caixas
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import ElementClickInterceptedException
from tkinter import messagebox
import time
import tkinter as tk
import os


#(<Inicio do Processo 1>)#################################################################################

print("==================================================================")
print(f"(1) Iniciando o processo")
print("==================================================================")
print()

# Função para medir o tempo de execução de cada etapa
def log_tempo(start, mensagem):
    print(f"{mensagem}: {time.time() - start:.2f} segundos")

# Configuração para acesso ao Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Caminho dinâmico para o credenciais.json
credenciais_path = os.path.join(os.path.dirname(__file__), 'credenciais.json')
if not os.path.exists(credenciais_path):
    print(f"Erro: Arquivo 'credenciais.json' não encontrado em {credenciais_path}.")
    print("Certifique-se de que o arquivo está no diretório correto.")
    exit(1)

# Autorização do Google Sheets
try:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credenciais_path, scope)
    client = gspread.authorize(credentials)
except Exception as e:
    print(f"Erro ao autorizar credenciais do Google Sheets: {e}")
    exit(1)

# URL do Google Sheets e nome da sheet
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing"
sheet_name = "CONTAORDEM"

try:
   
    sheet = client.open_by_url(spreadsheet_url).worksheet(sheet_name)
    print(f"Selecionando dados na sheet: {sheet_name}\n")
    data = sheet.get_all_records()
except Exception as e:
    print(f"Erro ao acessar a planilha ou a sheet: {e}")
    exit(1)

# Configuração do WebDriver
chrome_options = Options()
modo_headless = True  # Alterar para False durante depuração
if modo_headless:
    chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

webdriver_path = "G:\\O meu disco\\python\\chromedriver\\chromedriver.exe"
if not os.path.exists(webdriver_path):
    print(f"Erro: Caminho do ChromeDriver inválido: {webdriver_path}")
    exit(1)

# Inicializar o WebDriver
try:
     
    print("WebDriver inicializado com sucesso.")
    print()
except Exception as e:
    print("Erro ao iniciar o WebDriver:")
    print(f"- Verifique o caminho do ChromeDriver: {webdriver_path}")
    print(f"- Certifique-se de que a versão do ChromeDriver é compatível com o navegador.")
    print(f"Erro detalhado: {e}")
    exit(1)


# Função para encontrar elementos com timeout padronizado
def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        elemento = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
        #print(f"Elemento localizado com sucesso: {valor}")
        return elemento
    except Exception as e:
        print(f"Erro ao localizar o elemento: {valor}. Detalhes: {e}")
        driver.save_screenshot("erro.png")  # Salva o estado da página
        return None

#(<Fim do Processo 1/>)################################################################################# 


#(<Inicio do Processo 2>)#################################################################################

print("==================================================================")
print(f"(2) Iniciando o processo de acesso ao SOMA")
print("==================================================================")
print()

# URL do site para automação
url = 'https://verbodavida.info/apps/index.php'

print(f"Acessando o site: {url}")
driver.get(url)

# Iniciando o processo de login
print("Iniciando o processo de login...")

# Login no site
usuario = tentar_encontrar_elemento(driver, By.NAME, 'email')
if usuario:
    usuario.send_keys('familialopesemportugal@gmail.com')
    print("Campo 'email' preenchido com sucesso.")
else:
    print("Erro: Campo 'email' não encontrado!")

senha = tentar_encontrar_elemento(driver, By.NAME, 'senha')
if senha:
    senha.send_keys('P@1internet')
    print("Campo 'senha' preenchido com sucesso.")
else:
    print("Erro: Campo 'senha' não encontrado!")

botao_login = tentar_encontrar_elemento(driver, By.NAME, 'submit')
if botao_login:
    botao_login.click()
    print("Botão 'login' clicado com sucesso.")
    time.sleep(2)  # Tempo para garantir o carregamento da página seguinte
else:
    print("Erro: Botão 'login' não encontrado!")

# Tentativa de localizar e interagir com o botão SOMA
try:
    botao_soma = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, '285'))
    )
    botao_soma.click()
    print("Botão 'SOMA' clicado com sucesso!")
except TimeoutException:
    print("Erro: Botão 'SOMA' não encontrado ou não ficou disponível a tempo!")
    driver.save_screenshot("erro_soma.png")



for idx, linha in enumerate(data):
 
    # Aguarda o carregamento da página após o clique
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
        )
        print("Página carregada com sucesso após clicar no botão 'SOMA'.")
    except TimeoutException:
        #print("Erro: Página não carregou após clicar no botão 'SOMA'.")
        driver.save_screenshot("erro_pagina_soma.png")

#(<Fim do Processo 2/>)################################################################################# 


#(<Inicio do Processo 3>)#################################################################################

    print()
    print("==================================================================")
    print(f"(3) Iniciando os processos de interação com os dados antes dos inputs")
    print("==================================================================")
    print()

    # Carregar os dados da planilha e filtrar apenas as linhas onde 'DOC. SOMA' está vazio
    data = sheet.get_all_records()  # Obter os dados da planilha como uma lista de dicionários

    # Filtrar as linhas com 'DOC. SOMA' vazio ou com o valor "Em processamento".
    linhas_filtradas = [
        (index, row)
        for index, row in enumerate(data, start=2)
        if str(row.get('DOC. SOMA', '')).strip() in ['', 'Em processamento']
    ]

    # Processar as linhas filtradas
    if linhas_filtradas:
        for index, row in linhas_filtradas:
            # Escrever "Em processamento" na coluna 'DOC. SOMA' da linha atual
            cabecalho = sheet.row_values(1)  # Obter cabeçalho da planilha
            if 'DOC. SOMA' in cabecalho:
                coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1  # Obter índice 1-based da coluna
                try:
                    sheet.update_cell(index, coluna_doc_soma, "Em processamento")
                    print(f"Marcado 'Em processamento' na linha {index}, coluna 'DOC. SOMA'.")
                except Exception as e:
                    print(f"Erro ao atualizar a célula na planilha: {e}")
            else:
                print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
                continue
    else:
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA').")
        atualizar_caixas()

    #(<Fim do Processo 3/>)################################################################################# 
 

    # Imprimir os nomes das colunas disponíveis e o total de linhas filtradas para depuração
    if linhas_filtradas:
        
        #print(f"Colunas disponíveis na planilha: {data[0].keys()}")
        print(f"Total de linhas filtradas (DOC. SOMA = vazio): ({len(linhas_filtradas)})")
        
    else:
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA'")
        exit()  # Interrompe a execução do script


    # URL inicial
    url_inicial = "https://verbodavida.info/IVV/"

    # Função para redirecionar para a página inicial e validar o carregamento
    def redirecionar_para_inicio(driver, url, timeout=20):
        print(f"Redirecionando para a página inicial: {url}")
        driver.get(url)
        try:
            # Validar se a página inicial carregou corretamente
            WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
            )
            print("Página inicial carregada com sucesso!")
        except TimeoutException:
            print("Erro: Não foi possível carregar a página inicial.")
            driver.save_screenshot("erro_pagina_inicial.png")  # Salva uma captura de tela para análise
            raise

        
    # Iterar sobre as linhas filtradas
    for index, row in linhas_filtradas:
        print( )
        print(f"Processando a linha {index}: {row}")
        
        try:
            plano_de_conta_xl = row['PLANO DE CONTA']
            centro_custo_xl = row['CENTRO DE CUSTO']
            descricao_xl = row['DESCRIÇÃO SOMA']
            data_xl = row['DATA MOV.']
            valor_xl = row['IMPORTÂNCIA']
            forma_de_pagamento_xl = row['FORMA DE PAGAMENTO']
            caixa_xl = row['CAIXA']
            caixa_saida_xl = row['CAIXA SAIDA']
            tipo_xl = row['TIPO']
        except KeyError as e:
            print(f"Erro: A coluna '{e.args[0]}' não foi encontrada nesta linha. Dados ignorados.")
            continue
            
           

    if tipo_xl == "Entrada":
        atualizar_entradas()

        """
        elif tipo_xl == "Saídas":
            atualizar_saidas()

        elif tipo_xl == "Transferência":
            atualizar_transferencias()
            """

        #(<Fim do Processo 3/>)################################################################################# 
        
    #-------------------------------------------------------------------------------------


# Finalizar o WebDriver
try:
    driver.quit()
    print("WebDriver finalizado com sucesso.")
except Exception as e:
    print(f"Erro ao finalizar o WebDriver: {e}")

# Exibir mensagem de popup ao final do processamento
def mostrar_popup():
    import tkinter as tk
    from tkinter import messagebox

    try:
        # Criar a janela principal do Tkinter
        root = tk.Tk()
        root.withdraw()  # Ocultar a janela principal

        # Exibir o popup com a mensagem de conclusão
        messagebox.showinfo("Processamento Concluído", "O processamento foi finalizado com sucesso.")
        root.quit()
    except Exception as e:
        print(f"Erro ao exibir popup: {e}")

# Chamar a função para exibir o popup
mostrar_popup()

# Encerrar o programa
exit(0)
