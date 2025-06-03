import gspread
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
import time

#(1)#################################################################################

start_time = time.time()
print("(1)===================================================")
print(f"- Iniciando o processo")
print(" ")


# Função para medir o tempo de execução de cada etapa
def log_tempo(start, mensagem):
    print(f"{mensagem}: {time.time() - start:.2f} segundos")

# Configuração para acesso ao Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)

client = gspread.authorize(credentials)

# URL do Google Sheets e nome da sheet
spreadsheet_url = "https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing"
#sheet_name = "TESTE"
sheet_name = "CONTAORDEM"
print(f"Selecionando dados na sheet: {sheet_name}")
print(" ")


sheet = client.open_by_url(spreadsheet_url).worksheet(sheet_name)
data = sheet.get_all_records()  # Obter os dados da planilha como uma lista de dicionários

# Configuração do WebDriver para o Chrome com modo headless ------------------------

# Trecho para executar o processamento em background 
chrome_options = Options()

#"""
chrome_options.add_argument("--headless")  # Ativar o modo headless
chrome_options.add_argument("--disable-gpu")  # Otimizar para headless
chrome_options.add_argument("--no-sandbox")  # Necessário em alguns sistemas
chrome_options.add_argument("--disable-dev-shm-usage")  # Melhor para memória compartilhada
chrome_options.add_argument("--window-size=1920,1080")  # Tamanho da janela para evitar problemas de layout

#"""

webdriver_path = "G:\\O meu disco\\python\\chromedriver\\chromedriver.exe"  # Caminho atualizado para o ChromeDriver
service = Service(webdriver_path)
#----------------------------------------------------------------------------------- 

# Inicializar o WebDriver no modo headless

try:
    driver = webdriver.Chrome(service=service, options=chrome_options)
except Exception as e:
    print(f"Erro ao iniciar o WebDriver: {e}")
    exit(1)  # Finalizar o script em caso de falha
print( )
log_tempo(start_time, "Tempo para Iniciando o processo")
#----------------------------------------------------------------------------------------------


#(2)#################################################################################

start_time = time.time()
print(" ")
print("(2)===================================================")
print(f"Iniciando o processo de acesso ao SOMA")
print(" ")

# URL do site para automação
url = 'https://verbodavida.info/apps/index.php'

driver.get(url)

# Função para encontrar elementos com timeout e capturar o estado da página em caso de erro
def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
    except:
        print(f"Elemento não encontrado ou visível: {valor}")
        driver.save_screenshot("erro.png")  # Salva o estado da página
        return None

# Login no site
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

    time.sleep(2)
print( )
log_tempo(start_time, "Tempo para acessar o SOMA")
   

#(3)##################################################################################

start_time = time.time()
print(" ")
print("(3)===================================================")
print(f"Iniciando os processos de interração com a página antes dos inputs")
print(" ")

# Tentativa de localizar e interagir com o botão SOMA
try:
    botao_soma = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, '285'))
    )
    botao_soma.click()
    print("Botão 'SOMA' clicado com sucesso!")
except TimeoutException:
    print("Erro: Botão 'SOMA' não encontrado ou não ficou disponível a tempo!")

for idx, linha in enumerate(data):
    print("Processando todas as linhas!")
            
    # Aguarda o carregamento da página após o clique
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
        )
        print("Página carregada com sucesso após clicar no botão 'SOMA'.")
    except TimeoutException:
        print("Erro: Página não carregou após clicar no botão 'SOMA'.")

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
            print(f"Processando a linha {index}: {row}")

            # Escrever "Em processamento" na coluna 'DOC. SOMA' da linha atual
            cabecalho = sheet.row_values(1)  # Obter cabeçalho da planilha
            if 'DOC. SOMA' in cabecalho:
                coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1  # Obter índice 1-based da coluna
                sheet.update_cell(index, coluna_doc_soma, "Em processamento")
                print(f"Marcado 'Em processamento' na linha {index}, coluna 'DOC. SOMA'.")
            else:
                print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
                continue
    else:
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA').")

        # (1) Início do processo
        start_time = time.time()
        print("(1)===================================================")
        print("- Iniciando o processo de automação")
        print(" ")

        # Função para medir o tempo de execução de cada etapa
        def log_tempo(start, mensagem):
            print(f"{mensagem}: {time.time() - start:.2f} segundos")

        # Configuração para acesso ao Google Sheets
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
        client = gspread.authorize(credentials)

        # URL do Google Sheets e nome da sheet
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing"
        sheet_name = "Gerenciar Caixas"
        print(f"Selecionando dados na sheet: {sheet_name}")
        print(" ")

        sheet = client.open_by_url(spreadsheet_url).worksheet(sheet_name)
        data = sheet.get_all_records()  # Obter os dados da planilha como uma lista de dicionários

        # Obter o cabeçalho da sheet (linha 1) e seus índices
        cabecalho = sheet.row_values(1)
        indices_colunas = {coluna: idx + 1 for idx, coluna in enumerate(cabecalho)}  # Índices são 1-based no Google Sheets
        print("Índice e nomes das colunas carregados com sucesso:")
        for coluna, idx in indices_colunas.items():
            print(f" - {coluna}: índice {idx}")

        log_tempo(start_time, "Tempo para carregar a sheet e o cabeçalho")
        print(" ")

        # Configuração do WebDriver para o Chrome com modo headless
        chrome_options = Options()

        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--window-size=1920,1080")


        webdriver_path = "G:\\O meu disco\\python\\chromedriver\\chromedriver.exe"
        service = Service(webdriver_path)

        # Inicializar o WebDriver
        try:
            driver = webdriver.Chrome(service=service, options=chrome_options)
            print("WebDriver iniciado com sucesso.")
        except Exception as e:
            print(f"Erro ao iniciar o WebDriver: {e}")
            exit(1)

        # (2) Acessar o site e realizar o login
        start_time = time.time()
        print("(2)===================================================")
        print("Acessando o site e realizando o login")
        print(" ")

        url = 'https://verbodavida.info/apps/index.php'
        driver.get(url)

        # Função para encontrar elementos com timeout
        def tentar_encontrar_elemento(driver, by, valor, timeout=20):
            try:
                return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
            except TimeoutException:
                print(f"Erro: Elemento não encontrado: {valor}")
                driver.save_screenshot("erro_elemento.png")
                return None

        usuario = tentar_encontrar_elemento(driver, By.NAME, 'email')
        if usuario:
            usuario.send_keys('familialopesemportugal@gmail.com')
            print("Email preenchido com sucesso.")

        senha = tentar_encontrar_elemento(driver, By.NAME, 'senha')
        if senha:
            senha.send_keys('P@1internet')
            print("Senha preenchida com sucesso.")

        botao_login = tentar_encontrar_elemento(driver, By.NAME, 'submit')
        if botao_login:
            botao_login.click()
            print("Login realizado com sucesso.")

        log_tempo(start_time, "Tempo para acessar o site e realizar login")
        print(" ")

        #(3)##################################################################################

        start_time = time.time()
        print(" ")
        print("(3)===================================================")
        print(f"Iniciando os processos de interração com a página antes dos inputs")
        print(" ")

        # Tentativa de localizar e interagir com o botão SOMA
        try:
            botao_soma = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, '285'))
            )
            botao_soma.click()
            print("Botão 'SOMA' clicado com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'SOMA' não encontrado ou não ficou disponível a tempo!")

            """
            for idx, linha in enumerate(data):
                print("Processando todas as linhas!")
            """            
            # Aguarda o carregamento da página após o clique
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
                )
                print("Página carregada com sucesso após clicar no botão 'SOMA'.")
            except TimeoutException:
                print("Erro: Página não carregou após clicar no botão 'SOMA'.")

        # (4) Navegar para "Caixas/bancos"
        start_time = time.time()
        print("(4)===================================================")
        print("Navegando para a seção 'Caixas/bancos'")
        print(" ")

        try:
            botao_caixas_bancos = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div/div[1]/div'))
            )
            botao_caixas_bancos.click()
            print("Clicou na opção 'Caixas/bancos' com sucesso!")
        except TimeoutException:
            print("Erro: Elemento 'Caixas/bancos' não encontrado ou não ficou clicável a tempo!")

        log_tempo(start_time, "Tempo para acessar a seção 'Caixas/bancos'")
        print(" ")

        # (4) Processar as colunas específicas e informações correspondentes
        print("(4)===================================================")
        print("Processando colunas específicas")
        print(" ")

        # Recuperar valores específicos da primeira linha
        linha_referencia = data[0]
        caixa_banco_xl = linha_referencia.get('CAIXA BANCO', '')
        caixa_diario_xl = linha_referencia.get('CAIXA DIÁRIO', '')
        conta_banco_xl = linha_referencia.get('CAIXA ECONÔMICA MONTEPIO GERAL - CC', '')
        caixa_criancas_xl = linha_referencia.get('D. CRIANÇAS', '')
        caixa_livraria_xl = linha_referencia.get('D. LIVRARIA', '')
        caixa_cafe_xl = linha_referencia.get('VERBO CAFE', '')

        print("Valores carregados:")
        print(f" - CAIXA BANCO: {caixa_banco_xl}")
        print(f" - CAIXA DIÁRIO: {caixa_diario_xl}")
        print(f" - CONTA BANCO: {conta_banco_xl}")
        print(f" - CAIXA CRIANÇAS: {caixa_criancas_xl}")
        print(f" - CAIXA LIVRARIA: {caixa_livraria_xl}")
        print(f" - CAIXA CAFE: {caixa_cafe_xl}")


        # Obter o valor do caixa diário
        texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div[1]/div/div/div[2]/div[1]/span[2]')
        if texto_elemento:
            caixa_diario = texto_elemento.text.strip()  # Captura o texto do elemento
            print(f"Valor extraído para 'CAIXA DIÁRIO': {caixa_diario}")
            
            try:
                # Identificar o índice da coluna 'CAIXA DIÁRIO' com base no cabeçalho
                cabecalho = sheet.row_values(1)  # Obtém a linha do cabeçalho como lista
                if 'CAIXA DIÁRIO' in cabecalho:
                    coluna_caixa_diario = cabecalho.index('CAIXA DIÁRIO') + 1  # Índice 1-based

                    # Atualizar o valor na célula correspondente
                    linha_para_atualizar = 2  # Definir a linha que será atualizada (ajustar conforme necessário)
                    sheet.update_cell(linha_para_atualizar, coluna_caixa_diario, caixa_diario)
                    print(f"Valor '{caixa_diario}' atualizado na sheet na linha {linha_para_atualizar}, coluna 'CAIXA DIÁRIO'.")
                else:
                    print("Erro: Coluna 'CAIXA DIÁRIO' não encontrada no cabeçalho da planilha.")
            except ValueError as e:
                print(f"Erro ao encontrar a coluna: {e}")
            except Exception as e:
                print(f"Erro ao atualizar a sheet: {e}")
        else:
            print("Erro: Não foi possível extrair o valor para 'CAIXA DIÁRIO'.")

        # Obter o valor do caixa banco
        texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div[1]/span[2]')
        if texto_elemento:
            caixa_banco = texto_elemento.text.strip()  # Captura o texto do elemento
            print(f"Valor extraído para 'CAIXA BANCO': {caixa_banco}")
            
            try:
                # Identificar o índice da coluna 'CAIXA BANCO' com base no cabeçalho
                cabecalho = sheet.row_values(1)  # Obtém a linha do cabeçalho como lista
                if 'CAIXA BANCO' in cabecalho:
                    coluna_caixa_banco = cabecalho.index('CAIXA BANCO') + 1  # Índice 1-based

                    # Atualizar o valor na célula correspondente
                    linha_para_atualizar = 2  # Definir a linha que será atualizada (ajustar conforme necessário)
                    sheet.update_cell(linha_para_atualizar, coluna_caixa_banco, caixa_banco)
                    print(f"Valor '{caixa_banco}' atualizado na sheet na linha {linha_para_atualizar}, coluna 'CAIXA BANCO'.")
                else:
                    print("Erro: Coluna 'CAIXA BANCO' não encontrada no cabeçalho da planilha.")
            except ValueError as e:
                print(f"Erro ao encontrar a coluna: {e}")
            except Exception as e:
                print(f"Erro ao atualizar a sheet: {e}")
        else:
            print("Erro: Não foi possível extrair o valor para 'CAIXA BANCO'.")


        # Obter o valor do caixa Crianças
        texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div[3]/div/div/div[2]/div[1]/span[2]')
        if texto_elemento:
            caixa_criancas = texto_elemento.text.strip()  # Captura o texto do elemento
            print(f"Valor extraído para 'D. CRIANÇAS': {caixa_criancas}")
            
            try:
                # Identificar o índice da coluna 'D. CRIANÇAS' com base no cabeçalho
                cabecalho = sheet.row_values(1)  # Obtém a linha do cabeçalho como lista
                if 'D. CRIANÇAS' in cabecalho:
                    coluna_caixa_criancas = cabecalho.index('D. CRIANÇAS') + 1  # Índice 1-based

                    # Atualizar o valor na célula correspondente
                    linha_para_atualizar = 2  # Definir a linha que será atualizada (ajustar conforme necessário)
                    sheet.update_cell(linha_para_atualizar, coluna_caixa_criancas, caixa_criancas)
                    print(f"Valor '{caixa_criancas}' atualizado na sheet na linha {linha_para_atualizar}, coluna 'D. CRIANÇAS'.")
                else:
                    print("Erro: Coluna 'D. CRIANÇAS' não encontrada no cabeçalho da planilha.")
            except ValueError as e:
                print(f"Erro ao encontrar a coluna: {e}")
            except Exception as e:
                print(f"Erro ao atualizar a sheet: {e}")
        else:
            print("Erro: Não foi possível extrair o valor para 'D. CRIANÇAS'.")


        # Obter o valor do caixa Livraria
        texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div[4]/div/div/div[2]/div[1]/span[2]')
        if texto_elemento:
            caixa_livraria = texto_elemento.text.strip()  # Captura o texto do elemento
            print(f"Valor extraído para 'D. LIVRARIA': {caixa_livraria}")
            
            try:
                # Identificar o índice da coluna 'D. LIVRARIA' com base no cabeçalho
                cabecalho = sheet.row_values(1)  # Obtém a linha do cabeçalho como lista
                if 'D. LIVRARIA' in cabecalho:
                    coluna_caixa_livraria = cabecalho.index('D. LIVRARIA') + 1  # Índice 1-based

                    # Atualizar o valor na célula correspondente
                    linha_para_atualizar = 2  # Definir a linha que será atualizada (ajustar conforme necessário)
                    sheet.update_cell(linha_para_atualizar, coluna_caixa_livraria, caixa_livraria)
                    print(f"Valor '{caixa_livraria}' atualizado na sheet na linha {linha_para_atualizar}, coluna 'D. LIVRARIA'.")
                else:
                    print("Erro: Coluna 'D. LIVRARIA' não encontrada no cabeçalho da planilha.")
            except ValueError as e:
                print(f"Erro ao encontrar a coluna: {e}")
            except Exception as e:
                print(f"Erro ao atualizar a sheet: {e}")
        else:
            print("Erro: Não foi possível extrair o valor para 'D. LIVRARIA'.")

        exit()  # Interrompe a execução do script

    # Imprimir os nomes das colunas disponíveis e o total de linhas filtradas para depuração
    if linhas_filtradas:
        print( )
        print(f"Colunas disponíveis na planilha: {data[0].keys()}")
        print( )
        print(f"Total de linhas filtradas (DOC. SOMA = vazio): {len(linhas_filtradas)}")
        
    else:
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA'")
        exit()  # Interrompe a execução do script

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

    #-------------------------------------------------------------

    if tipo_xl != "Transferência":

        # Localizar e clicar na opção "Entradas/Saídas"
        try:
            elemento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o elemento esteja clicável
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
            )
            elemento.click()
            print( )
            print("Clicou na opção 'Entradas/Saídas' com sucesso!")
        except TimeoutException:
            print("Erro: Elemento 'Entradas/Saídas' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)

        # Após carregar a página de Entradas/Saídas, clicar no botão "Nova Entrada/Saída"
        try:
            nova_entrada_saida = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão esteja clicável
                EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "btn btn-primary") and contains(., "Nova Entrada/Saída")]'))
            )
            nova_entrada_saida.click()
            print("Clicou no botão 'Nova Entrada/Saída' com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Nova Entrada/Saída' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)

    #-------------------------------------------------------------
    
    print( )
    log_tempo(start_time, "Tempo para Iniciando os processos de interração") 

    #(4)#################################################################################

    start_time = time.time()
    print(" ")
    print("(4)===================================================")
    print(f"- O processo inicializado é para o tipo: {tipo_xl}")
   
    #-------------------------------------------------------------------------------------

    if tipo_xl == "Entrada":
        print(" ")
        print("🔥🔥🔥🔥🔥🔥🔥🔥 O tipo do documento é 'Entrada'. Continuando o processamento...")
        print(" ")

        #-------------------------------------------------------------------------------------
            
        # Preencher os campos na página "Entrada"
            
        botao_radio_entrada = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[2]/input')
        botao_radio_entrada.click()
        print("Selecionado o botão de rádio 'Entrada'")
        
        time.sleep(1)
        print( )
        

    
        #4.1#################################################################################
        start_time = time.time()
        print("4.1===================================================")
        print(f"Iniciando o processo de input de dados para a linha {index} e para o tipo {tipo_xl}")
        print(" ")

        # Preencher o campo "Plano de Conta"
        plano_de_conta = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/span/span[1]/span/span[1]')
        if plano_de_conta:
            plano_de_conta.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_plano_de_conta = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_plano_de_conta.clear()  # Limpar qualquer valor pré-existente
            selecao_plano_de_conta.send_keys(plano_de_conta_xl)  # Inserir o valor do plano de conta
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), plano_de_conta_xl
                )
            )
            selecao_plano_de_conta.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"Plano de conta preenchido com sucesso: {plano_de_conta_xl}")
        else:
            print("Erro: Campo para selecionar plano de conta não encontrado.")
            time.sleep(2)


        # Preencher o campo "Centro de custo"
        centro_de_custo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span')
        if centro_de_custo:
            centro_de_custo.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_centro_de_custo = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_centro_de_custo.clear()  # Limpar qualquer valor pré-existente
            selecao_centro_de_custo.send_keys(centro_custo_xl)  # Inserir o valor do centro de custo
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), centro_custo_xl
                )
            )
            selecao_centro_de_custo.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"Centro de custo preenchido com sucesso: {centro_custo_xl}")
        else:
            print("Erro: Campo para selecionar Centro de custo não encontrado.")
            time.sleep(2)

        
        # Selecionar o campo "Descrição"
        descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/input')
        if descricao_campo:
            descricao_campo.send_keys(descricao_xl)
            print(f"Descrição preenchida com sucesso: {descricao_xl}")
        else:
            print("Campo de descrição não encontrado!")
            time.sleep(2)
                
        # Selecionar o campo "Data"
        data_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[11]/div/div/input')
        if data_campo:
            data_campo.send_keys(data_xl)
            print(f"Data preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data não encontrado!")
            time.sleep(2)
        
        # Selecionar o campo "Valor"
        valor_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[12]/div/input')
        if valor_campo:
            valor_campo.send_keys(str(valor_xl))
            time.sleep(1)  # Aguardar para garantir que o valor seja processado
            data_campo.send_keys(Keys.ENTER)  # Pressionar Enter para confirmar o valor
            print(f"Valor preenchido com sucesso: {valor_xl}")
        else:
            print("Campo de valor não encontrado!")
            time.sleep(2)
                
        # Selecionar o campo "Forma de pagamento"
        forma_de_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[16]/div/span/span[1]/span')
        if forma_de_pagamento:
            forma_de_pagamento.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_forma_de_pagamento = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_forma_de_pagamento.clear()  # Limpar qualquer valor pré-existente
            selecao_forma_de_pagamento.send_keys(forma_de_pagamento_xl)  # Inserir o valor da forma de pagamento
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), forma_de_pagamento_xl
                )
            )
            selecao_forma_de_pagamento.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"Forma de pagamento preenchida com sucesso: {forma_de_pagamento_xl}")
        else:
            print("Erro: Campo para selecionar Forma de pagamento não encontrado.")
            time.sleep(5)
        
        
        if forma_de_pagamento_xl == "DINHEIRO":
            print(f"A forma de pagamento é {forma_de_pagamento_xl}. Selecionar o campo Caixa diário.")
        
            # Selecionar o campo "Caixa Entrada"
            caixa_entrada_dinheiro = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[20]/div')
            if caixa_entrada_dinheiro:
                caixa_entrada_dinheiro.click()  # Clica no campo para abrir o dropdown
                print("Dropdown 'Caixa Entrada' aberto com sucesso.")

                try:
                    
                    time.sleep(5)
                    caixa_entrada_dinheiro_select = caixa_entrada_dinheiro.find_element(By.TAG_NAME, "select")
                    opcoes_de_caixa = caixa_entrada_dinheiro_select.find_elements(By.TAG_NAME, "option")

                    for opcao in opcoes_de_caixa:
                        texto_da_opcao = opcao.text
                        if texto_da_opcao in caixa_xl:
                            opcao.click()
                            print(f"Sugestão correspondente '{caixa_xl}' selecionada com sucesso.")

                except TimeoutException:
                    print("Erro: Não foi possível encontrar o valor no dropdown dentro do tempo limite.")
                except Exception as e:
                    print(f"Erro inesperado ao selecionar o valor no dropdown: {e}")
            else:
                print("Erro: Campo para selecionar 'Caixa Entrada' não encontrado.")
                time.sleep(1)
        
        else:
                print(f"A forma de pagamento é diferente de DINHEIRO. pular a lógia de processamento...")

        
                
        print( )
        if forma_de_pagamento_xl == "TRANSFERÊNCIA BANCÁRIA":
            print(f"A forma de pagamento é {forma_de_pagamento_xl}. Selecionar o campo Caixa banco.")

            # Selecionar o campo "Caixa" - (Saída)
            caixa_entrada_TB = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[20]/div')
            if caixa_entrada_TB:
                caixa_entrada_TB.click()  # Clica no campo para abrir o dropdown
                print("Dropdown 'Caixa Entrada' aberto com sucesso.")
                
                try:
                    # Criar um objeto Select para manipular o elemento <select>
                    caixa_entrada_TB = Select(caixa_xl)

                    # Selecionar a opção pelo valor (assumindo que forma_de_pagamento_xl contém o texto correspondente)
                    caixa_entrada_TB.select_by_index(3)
                    #caixa_entrada_TB.click()
                    print(f"Forma de pagamento selecionada com sucesso: {caixa_entrada_TB}")

                except Exception as e:
                    print( )
                    print("🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥🔥")
                    print(f"Erro ao selecionar a forma de pagamento: {e}")
                    
            else:
                print("Campo 'Forma de pagamento' não encontrado!")
            time.sleep(2)

        else:
                print(f"A forma de pagamento é diferente de {forma_de_pagamento_xl}. pular a lógia de processamento...")
        print( )
     
     
        # Selecionar o campo "Obs"
        observacao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[21]/div/textarea')
        if observacao_campo:
            observacao_campo.click()
            observacao_campo.send_keys(descricao_xl)
            print(f"Descrição preenchida com sucesso: {descricao_xl}")
            
        else:
            print("Campo de descrição não encontrado!")
            time.sleep(2)
        

        # Salvar o formulário
        botao_salvar = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[30]/div/button[1]')
        if botao_salvar:
            botao_salvar.click()  # Clicar no botão salvar
            print("Formulário salvo com sucesso!")
            time.sleep(5)  # Aguarda para garantir que o formulário seja processado antes de continuar
        else:
            print("Botão de salvar o formulário não encontrado!")
            time.sleep(2)
        print( )
        log_tempo(start_time, "Tempo para interragir com os processos de input das informações")
        
        

        #4.2#################################################################################
        start_time = time.time()
        print(" ")
        print("4.2===================================================")
        print(f"Iniciando o processo de baixa do documento")
        print(" ")

        if forma_de_pagamento_xl == "TRANSFERÊNCIA BANCÁRIA":
            print("A forma de pagamento é TRANSFERÊNCIA BANCÁRIA. iniciar processo de baixa de pagamento...")
            print(" ")

            # Realizar pagamento
            botao_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/ul/li/div/div[2]/button[1]')
            if botao_pagamento:
                botao_pagamento.click()  # Clicar no botão salvar
                print("Realizar pagamento salvo com sucesso!")
                time.sleep(2)  # Aguarda para garantir que o formulário seja processado antes de continuar
            else:
                print("Botão de Realizar pagamento não encontrado!")
                
            # Inserir baixa
            botao_inserir_baixa = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button')
            if botao_inserir_baixa:
                botao_inserir_baixa.click()  # Clicar no botão salvar
                print("Inserir Baixa salvo com sucesso!")
                time.sleep(2)  # Aguarda para garantir que o formulário seja processado antes de continuar
            else:
                print("Botão de Inserir Baixa não encontrado!")
                
            
            # Selecionar o campo "Data"
            data_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input')
            if data_pagamento:
                data_pagamento.clear()  # Limpar o campo antes de inserir nova data
                data_pagamento.send_keys(data_xl)  # Inserir a nova data
                print(f"Data início preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data inicial não encontrado!")
            time.sleep(1)

            # Clicar no pop-up "Data"
            observacao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[2]')
            if observacao_campo:
                observacao_campo.click()
                print(f"Click na janela pop-up com sucesso")
                time.sleep(2)
            else:
                print("Campo de descrição não encontrado!")
                

            # Salvar baixa
            
            try:
                botao_salvar_baixa = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[3]/button'))
                )
                botao_salvar_baixa.click()  # Clicar no botão salvar
                print("Salvar Baixa salvo com sucesso!")
                time.sleep(2)
            except TimeoutException:
                print("Erro: Botão 'Salvar Baixa' não encontrado ou não ficou clicável a tempo!")
                
                
            # Clicar no botão "OK Baixa"
            try:
                botao_OK_baixa = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
                )
                botao_OK_baixa.click()  # Clicar no botão "OK Baixa"
                print("Botão 'OK Baixa' clicado com sucesso!")
                time.sleep(2)
            except TimeoutException:
                print("Erro: Botão 'OK Baixa' não encontrado ou não ficou clicável a tempo!")
                
            print( )
            log_tempo(start_time, "Tempo para realizar a compensação do documento")
        else:
            print("Forma de pagamento é diferente de 'TRANSFERÊNCIA BANCÁRIA'. Pulando o processo de baixa de pagamento.")
    

        #4.3#################################################################################
        start_time = time.time()    
        print(" ")
        print("4.3===================================================")
        print(f"Iniciando o processo pesquisa do nº soma do documento")
        print(" ")

        # Clicar botão voltar a sessão "Entradas/Saídas"
        botao_voltar = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]')
        if botao_voltar:
            botao_voltar.click() # Clicar no botão salvar
            print("Clicou na opção 'Entradas/Saídas' com sucesso!")
            time.sleep(1)  # Aguarda para garantir que o formulário seja processado antes de continuar
        else:
            print("Elemento 'Entradas/Saídas' não encontrado!")
            time.sleep(2)
            
        # Selecionar o campo "Pesquisar"
        pesquisar_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/input')
        if pesquisar_campo:
            pesquisar_campo.send_keys(descricao_xl)
            print(f"Campo pesquisar a descrição preenchida com sucesso: {descricao_xl}")
            
        else:
            print("Campo de pesquisar a descrição não encontrado!")
        time.sleep(1)
            
        # Clicar botão "periodo"
        
        botao_radio_periodo = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/div/input')
        botao_radio_periodo.click()
        print("Selecionado o botão de rádio 'Periodo'")
        time.sleep(1)


        # Preencher os campos na página "Data de Pagamento"
        
        botao_radio_data_pagamento = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[8]/div/div[2]/input')
        botao_radio_data_pagamento.click()
        print("Selecionado o botão de rádio 'Data de Pagamento'")
        time.sleep(1)

        # Selecionar intervalo de "Data"
        data_intervaloA = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[1]/div/input')
        if data_intervaloA:
            data_intervaloA.clear()  # Limpar o campo antes de inserir nova data
            data_intervaloA.send_keys(data_xl)  # Inserir a nova data
            print(f"Data início preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data inicial não encontrado!")
        time.sleep(1)

        data_intervaloB = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[3]/div/input')
        if data_intervaloB:
            data_intervaloB.clear()  # Limpar o campo antes de inserir nova data
            data_intervaloB.send_keys(data_xl)  # Inserir a nova data
            print(f"Data fim preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data final não encontrado!")
        time.sleep(1)

        # Após carregar a página de Entradas/Saídas, clicar no botão "Pesquisar"
        
        try:
            pesquisar_documento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/button'))
            )
            pesquisar_documento.click()  # Clicar no botão "Pesquisar"
            print("Botão 'Pesquisar' clicado com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Pesquisar' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)
    
        # Obter o número do documento gerado após salvar
        texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/div/div/table/tbody/tr/td[3]')
        if texto_elemento:
            doc_soma = texto_elemento.text.strip()  # Captura o texto do elemento
            print(f"Número do documento extraído: {doc_soma}")

            # Atualizar o valor na planilha após processar a linha
            cabecalho = sheet.row_values(1)  # Pega o cabeçalho da primeira linha
            if 'DOC. SOMA' in cabecalho:
                coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1  # Índice 1-based para o Google Sheets
                sheet.update_cell(index, coluna_doc_soma, doc_soma)  # Atualiza a célula na linha correta
                print(f"Texto '{doc_soma}' gravado na linha {index}, coluna 'DOC. SOMA'.")
            else:
                print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
        else:
            print(f"Não foi possível extrair o número do documento para a linha {index}.")
            time.sleep(2)
        
        
               
        print(" ")
        print(f"Processamento finalizado para a linha {index}")
        print("===================================================")
        print( )
        log_tempo(start_time, "Tempo total para a pesquisar")
        print(" ")

    #-------------------------------------------------------------------------------------
    
    if tipo_xl == "Saída":
        print(" ")
        print("🔥🔥🔥🔥🔥🔥🔥🔥 O tipo do documento é 'Saída'. Continuando o processamento...")
        print(" ")
 
        # Preencher os campos na página "Saída"
        
        botao_radio_saida = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[1]/input')
        botao_radio_saida.click()
        print("Selecionado o botão para o processo de 'Saída'")
        time.sleep(1)
        
        #-------------------------------------------------------------------------------------

        start_time = time.time()
        print(" ")
        print("4.1===================================================")
        print(f"Iniciando o processo de input de dados para a linha {index} e o tipo {tipo_xl}")
        print(" ")

        # Preencher o campo "Plano de Conta"
        plano_de_conta = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/span/span[1]/span/span[1]')
        if plano_de_conta:
            plano_de_conta.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_plano_de_conta = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_plano_de_conta.clear()  # Limpar qualquer valor pré-existente
            selecao_plano_de_conta.send_keys(plano_de_conta_xl)  # Inserir o valor do plano de conta
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), plano_de_conta_xl
                )
            )
            selecao_plano_de_conta.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"Plano de conta preenchido com sucesso: {plano_de_conta_xl}")
        else:
            print("Erro: Campo para selecionar plano de conta não encontrado.")
            time.sleep(2)

        # Preencher o campo "Centro de custo"
        centro_de_custo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span')
        if centro_de_custo:
            centro_de_custo.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_centro_de_custo = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_centro_de_custo.clear()  # Limpar qualquer valor pré-existente
            selecao_centro_de_custo.send_keys(centro_custo_xl)  # Inserir o valor do centro de custo
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), centro_custo_xl
                )
            )
            selecao_centro_de_custo.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"Centro de custo preenchido com sucesso: {centro_custo_xl}")
        else:
            print("Erro: Campo para selecionar Centro de custo não encontrado.")
            time.sleep(2)
        
        # Selecionar o campo "Descrição"
        descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/input')
        if descricao_campo:
            descricao_campo.send_keys(descricao_xl)
            print(f"Descrição preenchida com sucesso: {descricao_xl}")
            
        else:
            print("Campo de descrição não encontrado!")
            time.sleep(2)
                
        # Selecionar o campo "Data"
        data_vencimento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/div/input')
        if data_vencimento:
            data_vencimento.send_keys(data_xl)
            print(f"Data vencimento preenchida com sucesso: {data_xl}")
        else:
            print("Campo Data vencimento não encontrado!")
            time.sleep(2)
        
        
        # Selecionar o campo "Valor"
        valor_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[12]/div/input')
        if valor_campo:
            valor_campo.send_keys(str(valor_xl))
            time.sleep(1)  # Aguardar para garantir que o valor seja processado
            valor_campo.send_keys(Keys.ENTER)  # Pressionar Enter para confirmar o valor
            print(f"Valor preenchido com sucesso: {valor_xl}")
        else:
            print("Campo de valor não encontrado!")
            time.sleep(2)
        
        
        # Selecionar o campo "Obs"
        observacao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[21]/div/textarea')
        if observacao_campo:
            observacao_campo.click()
            observacao_campo.send_keys(descricao_xl)
            print(f"Descrição preenchida com sucesso: {descricao_xl}")
            
        else:
            print("Campo de descrição não encontrado!")
            time.sleep(2)

        print( )
        log_tempo(start_time, "Tempo para o processo de input de dados") 

        #-------------------------------------------------------------------------------------
        
        start_time = time.time()
        print(" ")
        print("4.2===================================================")
        print(f"Iniciando o processo de inserir pagamento")
        print(" ")

        # Realizar pagamento
        botao_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/ul/li/div/div[2]/button[1]')
        if botao_pagamento:
            botao_pagamento.click()  # Clicar no botão salvar
            print("Realizar pagamento salvo com sucesso!")
            time.sleep(5)  # Aguarda para garantir que o formulário seja processado antes de continuar
        else:
            print("Botão de Realizar pagamento não encontrado!")
            time.sleep(2)
        
        # Inserir pagamento
        botao_inserir_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[29]/div/div[2]/a')
        if botao_inserir_pagamento:
            botao_inserir_pagamento.click()  # Clicar no botão salvar
            print("Inserir pagamento salvo com sucesso!")
            time.sleep(2)  # Aguarda para garantir que o formulário seja processado antes de continuar
        else:
            print("Botão de Inserir pagamento não encontrado!")
            time.sleep(2)
        
        # Selecionar o campo "Data"
        data_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input')
        if data_pagamento:
            data_pagamento.clear()  # Limpar o campo antes de inserir nova data
            data_pagamento.send_keys(data_xl)  # Inserir a nova data
            print(f"Data início preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data inicial não encontrado!")
        time.sleep(1)
        
        # Selecionar o campo "Forma de pagamento" - (Saída)
        forma_de_pagamento1 = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select')
        if forma_de_pagamento1:
            try:
                # Criar um objeto Select para manipular o elemento <select>
                select_forma_pagamento1 = Select(forma_de_pagamento1)

                # Selecionar a opção pelo valor (assumindo que forma_de_pagamento_xl contém o texto correspondente)
                select_forma_pagamento1.select_by_visible_text(forma_de_pagamento_xl)
                print(f"Forma de pagamento selecionada com sucesso: {forma_de_pagamento_xl}")
            except Exception as e:
                print(f"Erro ao selecionar a forma de pagamento: {e}")
        else:
            print("Campo 'Forma de pagamento' não encontrado!")
        time.sleep(2)

        # Selecionar o campo "Caixa" - (Saída)
        caixa_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[4]/div[2]/div/select')
        if caixa_pagamento:
            try:
                # Criar um objeto Select para manipular o elemento <select>
                caixa_pagamento = Select(caixa_pagamento)

                # Selecionar a opção pelo valor (assumindo que forma_de_pagamento_xl contém o texto correspondente)
                caixa_pagamento.select_by_visible_text(caixa_xl)
                print(f"Caixa para pagamento selecionada com sucesso: {caixa_xl}")
            except Exception as e:
                print(f"Erro ao selecionar a Caixa para pagamento: {e}")
        else:
            print("Campo 'Caixa para pagamento' não encontrado!")
        time.sleep(2)


        # Verificar se a forma de pagamento é "TRANSFERÊNCIA BANCÁRIA"-------------------------------------------------

        if forma_de_pagamento_xl == "TRANSFERÊNCIA BANCÁRIA":
            print("A forma de pagamento é TRANSFERÊNCIA BANCÁRIA. atualizar campo Nº Documento...")

            # Selecionar o campo "Nº Documento"
            numero_documento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[3]/div/input')
            if numero_documento:
                numero_documento.send_keys(descricao_xl)
                print(f"Número do documento preenchido com sucesso: {descricao_xl}")
            else:
                print("Campo 'Nº Documento' não encontrado!")
                time.sleep(2)
        else:
            print(f"A forma de pagamento não é 'TRANSFERÊNCIA BANCÁRIA' (valor atual: {forma_de_pagamento_xl}). Ignorando este processo.")

        
        # Clicar no botão "Salvar Pagamento" - (Saída)
        try:
            botao_Salvar_pagamento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[3]/button'))
            )
            botao_Salvar_pagamento.click()  # Clicar no botão "OK Baixa"
            print("Botão 'Salvar Pagamento' clicado com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Salvar Pagamento' não encontrado ou não ficou clicável a tempo!")   
        time.sleep(2)

        # Clicar no botão "OK pagamento"
        try:
            botao_OK_pagamento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
            )
            time.sleep(5)
            botao_OK_pagamento.click()  # Clicar no botão "OK Baixa"
            print("Botão 'OK Baixa' clicado com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'OK Baixa' não encontrado ou não ficou clicável a tempo!")
        time.sleep(2)

        print(" ")
        log_tempo(start_time, "Tempo para interragir com inserir pagamento")
        
        #-------------------------------------------------------------------------------------
        
        start_time = time.time()
        print(" ")
        print("4.3===================================================")
        print(f"Iniciando o processo de baixa do documento")
        print(" ")

        #----------------------------------------------------------------------------------------

        if forma_de_pagamento_xl == "TRANSFERÊNCIA BANCÁRIA":
            print("A forma de pagamento é TRANSFERÊNCIA BANCÁRIA. Continuando o processamento...")

        #----------------------------------------------------------------------------------------
                    
            # Inserir baixa
            botao_inserir_baixa = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button')
            if botao_inserir_baixa:
                botao_inserir_baixa.click()  # Clicar no botão salvar
                print("Inserir Baixa salvo com sucesso!")
                time.sleep(2)  # Aguarda para garantir que o formulário seja processado antes de continuar
            else:
                print("Botão de Inserir Baixa não encontrado!")
                    
                
            # Selecionar o campo "Data"
            data_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input')
            if data_pagamento:
                data_pagamento.clear()  # Limpar o campo antes de inserir nova data
                data_pagamento.send_keys(data_xl)  # Inserir a nova data
                print(f"Data início preenchida com sucesso: {data_xl}")
                time.sleep(2)
            else:
                print("Campo de data inicial não encontrado!")
                

            # Clicar no pop-up "Baixa"
            observacao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[3]')
            if observacao_campo:
                observacao_campo.click()
                print(f"Click na janela pop-up com sucesso")
                time.sleep(2)
            else:
                print("Campo de descrição não encontrado!")
            
            
            # Salvar baixa
                
            try:
                    botao_salvar_baixa = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[3]/button'))
                    )
                    botao_salvar_baixa.click()  # Clicar no botão salvar
                    print("Salvar Baixa salvo com sucesso!")
                    time.sleep(2)
            except TimeoutException:
                    print("Erro: Botão 'Salvar Baixa' não encontrado ou não ficou clicável a tempo!")
                    
                    
                    # Clicar no botão "OK Baixa"
            try:
                    botao_OK_baixa = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                        EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
                    )
                    botao_OK_baixa.click()  # Clicar no botão "OK Baixa"
                    print("Botão 'OK Baixa' clicado com sucesso!")
                    time.sleep(2)
            except TimeoutException:
                    print("Erro: Botão 'OK Baixa' não encontrado ou não ficou clicável a tempo!")
                    
                    print( )
                    log_tempo(start_time, "Tempo para realizar a compensação do documento")
            else:
                    print("Forma de pagamento é diferente de 'TRANSFERÊNCIA BANCÁRIA'. Pulando o processo de baixa do documento.")
                
            
        # Clicar botão voltar a sessão "Entradas/Saídas"
        botao_voltar = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]')
        if botao_voltar:
                botao_voltar.click() # Clicar no botão salvar
                print("Clicou na opção 'Entradas/Saídas' com sucesso!")
                time.sleep(1)  # Aguarda para garantir que o formulário seja processado antes de continuar
        else:
                print("Elemento 'Entradas/Saídas' não encontrado!")
                time.sleep(2)

        print(" ")
        log_tempo(start_time, "Tempo para interragir com baixa do documento")

        #-------------------------------------------------------------------------------------
        
        start_time = time.time()      
        print(" ")
        print("4.4===================================================")
        print(f"Iniciando o processo pesquisa do nº soma do documento")
        print(" ")

        # Selecionar o campo "Pesquisar"
        pesquisar_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/input')
        if pesquisar_campo:
                pesquisar_campo.send_keys(descricao_xl)
                print(f"Campo pesquisar a descrição preenchida com sucesso: {descricao_xl}")
                
        else:
                print("Campo de pesquisar a descrição não encontrado!")
        time.sleep(1)
            
        # Clicar botão "periodo"
        
        botao_radio_periodo = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/div/input')
        botao_radio_periodo.click()
        print("Selecionado o botão de rádio 'Periodo'")
        time.sleep(1)


        # Preencher os campos na página "Data de Pagamento"
        
        botao_radio_data_pagamento = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[8]/div/div[2]/input')
        botao_radio_data_pagamento.click()
        print("Selecionado o botão de rádio 'Data de Pagamento'")
        time.sleep(1)

        # Selecionar intervalo de "Data"
        data_intervaloA = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[1]/div/input')
        if data_intervaloA:
            data_intervaloA.clear()  # Limpar o campo antes de inserir nova data
            data_intervaloA.send_keys(data_xl)  # Inserir a nova data
            print(f"Data início preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data inicial não encontrado!")
        time.sleep(1)

        data_intervaloB = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[3]/div/input')
        if data_intervaloB:
            data_intervaloB.clear()  # Limpar o campo antes de inserir nova data
            data_intervaloB.send_keys(data_xl)  # Inserir a nova data
            print(f"Data fim preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data final não encontrado!")
        time.sleep(1)

        # Após carregar a página de Entradas/Saídas, clicar no botão "Pesquisar"
        
        try:
            pesquisar_documento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/button'))
            )
            pesquisar_documento.click()  # Clicar no botão "Pesquisar"
            print("Botão 'Pesquisar' clicado com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Pesquisar' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)
    
        # Obter o número do documento gerado após salvar
        texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/div/div/table/tbody/tr/td[3]')
        if texto_elemento:
            doc_soma = texto_elemento.text.strip()  # Captura o texto do elemento
            print(f"Número do documento extraído: {doc_soma}")

            # Atualizar o valor na planilha após processar a linha
            cabecalho = sheet.row_values(1)  # Pega o cabeçalho da primeira linha
            if 'DOC. SOMA' in cabecalho:
                coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1  # Índice 1-based para o Google Sheets
                sheet.update_cell(index, coluna_doc_soma, doc_soma)  # Atualiza a célula na linha correta
                print(f"Texto '{doc_soma}' gravado na linha {index}, coluna 'DOC. SOMA'.")
            else:
                print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
        else:
            print(f"Não foi possível extrair o número do documento para a linha {index}.")
            sheet.update_cell(index, coluna_doc_soma, "Em erro")
            print(f"Processamento concluído para a linha {index}.")
            time.sleep(2)
        print( )
        log_tempo(start_time, "Tempo total para a pesquisar")
    
        print(" ")
        print(f"Processamento finalizado para a linha {index}")
        print("4.5===================================================")
        
    #-------------------------------------------------------------------------------------      
               
    if tipo_xl == "Transferência":
        print(" ")
        print("🔥🔥🔥🔥🔥🔥🔥🔥 O tipo do documento é Transferência. Continuando o processamento...")
        print(" ")

                    
        # Localizar e clicar na opção "Transferências Caixas"
        try:
            caixas_bancos = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o elemento esteja clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div/div[4]/div/div/span'))
            )
            time.sleep(5)
            caixas_bancos.click()
            print("Clicou na opção 'caixas bancos' com sucesso!")
        except TimeoutException:
            print("Erro: Elemento 'caixas bancos' não encontrado ou não ficou clicável a tempo!")
        
        # Após carregar a página de Entradas/Saídas, clicar no botão "Nova Transferência"
        try:
            nova_transferencia = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão esteja clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/a'))
            )
            time.sleep(5)
            nova_transferencia.click()
            print("Clicou no botão 'Nova Entrada/Saída' com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Nova Entrada/Saída' não encontrado ou não ficou clicável a tempo!")

        log_tempo(start_time, "Tempo para carregar as páginas até a página transferência")
                 
        # Selecionar o campo "Caixa Saída" - (Transferência)
        caixa_saida = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[2]/div')
        if caixa_saida:
            time.sleep(5)
            caixa_saida.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_caixa_saida = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_caixa_saida.clear()  # Limpar qualquer valor pré-existente
            selecao_caixa_saida.send_keys(caixa_saida_xl)  # Inserir o valor do plano de conta
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), caixa_saida_xl
                )
            )
            
            selecao_caixa_saida.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"CAIXA DIÁRIO preenchido com sucesso: {caixa_saida_xl}")
        else:
            print("Erro: Campo para selecionar CAIXA DIÁRIO não encontrado.")
            
        time.sleep(1)  # Aguardar para garantir que o valor seja processado

        # Selecionar o campo "Valor"
        valor_transferencia = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[4]/div/input')
        if valor_transferencia:
            valor_transferencia.send_keys(str(valor_xl))
            time.sleep(5)  # Aguardar para garantir que o valor seja processado
            #valor_transferencia.send_keys(Keys.ENTER)  # Pressionar Enter para confirmar o valor
            print(f"Valor preenchido com sucesso: {valor_xl}")
        else:
            print("Campo de valor não encontrado!")

        time.sleep(1)  # Aguardar para garantir que o valor seja processado

        # Selecionar o campo "Caixa Entrada" - (Transferência)
        caixa_entrada = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[5]/div')
        if caixa_entrada:
            time.sleep(5)
            caixa_entrada.click()  # Clica no campo para abrir o dropdown
            
            # Localizar o campo de busca no dropdown
            selecao_caixa_entrada = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
            )
            
            # Preencher o valor e confirmar com ENTER
            selecao_caixa_entrada.clear()  # Limpar qualquer valor pré-existente
            selecao_caixa_entrada.send_keys(caixa_xl)  # Inserir o valor do plano de conta
            WebDriverWait(driver, 2).until(  # Aguarda até que a sugestão apareça
                EC.text_to_be_present_in_element_value(
                    (By.XPATH, '/html/body/span/span/span[1]/input'), caixa_xl
                )
            )
            
            selecao_caixa_entrada.send_keys(Keys.ENTER)  # Confirmar a seleção com ENTER
            print(f"CAIXA DIÁRIO preenchido com sucesso: {caixa_xl}")
        else:
            print("Erro: Campo para selecionar CAIXA DIÁRIO não encontrado.")

        time.sleep(1)  # Aguardar para garantir que o valor seja processado
        
        # Selecionar o campo "Data"
        data_trensferencia = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[8]/div/div/input')
        if data_trensferencia:
            data_trensferencia.clear()  # Limpar o campo antes de inserir nova data
            data_trensferencia.send_keys(data_xl)  # Inserir a nova data
            print(f"Data de transferência preenchida com sucesso: {data_xl}")
        else:
            print("Campo de data de transferencia não encontrado!")

        time.sleep(1)  # Aguardar para garantir que o valor seja processado

        # Selecionar o campo "Descrição"
        descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[9]/div/textarea')
        if descricao_campo:
            descricao_campo.send_keys(descricao_xl)
            print(f"Descrição preenchida com sucesso: {descricao_xl}")
            
        else:
            print("Campo de descrição não encontrado!")

        # Salvar transferência - (Saída)
        try:
            botao_salvar_transferencia = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[10]/div/button[1]'))
            )
            botao_salvar_transferencia.click()  # Clicar no botão salvar
            print("Salvar pagamento salvo com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Salvar pagamento' não encontrado ou não ficou clicável a tempo!")

        
        # Clicar no botão "OK A transferência foi salva" - (Saída)
        try:
            botao_OK_pagamento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o botão fique clicável
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
            )
            botao_OK_pagamento.click()  # Clicar no botão "OK Baixa"
            print("Botão OK:'A transferência foi salva' ")
        except TimeoutException:
            print("Erro: Botão 'OK transferencia' não encontrado ou não ficou clicável a tempo!")   

        # Clicar botão voltar a sessão "Entradas/Saídas"
        botao_voltar = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[10]/div/button[2]')
        if botao_voltar:
            botao_voltar.click() # Clicar no botão salvar
            print("Clicou na opção 'Entradas/Saídas' com sucesso!")
            time.sleep(1)  # Aguarda para garantir que o formulário seja processado antes de continuar
        else:
            print("Elemento 'Entradas/Saídas' não encontrado!")
        
        # Atualizar o valor na planilha após processar a linha
        cabecalho = sheet.row_values(1)  # Pega o cabeçalho da primeira linha
        if 'DOC. SOMA' in cabecalho:
            coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1  # Índice 1-based para o Google Sheets
            texto_fixo = "Transferido"  # Texto fixo a ser inserido
            sheet.update_cell(index, coluna_doc_soma, texto_fixo)  # Atualiza a célula na linha correta
            print(f"Texto fixo '{texto_fixo}' gravado na linha {index}, coluna 'DOC. SOMA'.")
        else:
            print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")

    

    #-------------------------------------------------------------------------------------
input("Pressione ENTER para sair...")
    
# Finalizar o WebDriver
driver.quit()

