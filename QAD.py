import gspread
import sys
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
from tkinter import messagebox
import time
import tkinter as tk
import os
import threading

#(1)#################################################################################
print()
print("==================================================================")
print(f"(1) Iniciando o processo")
print("==================================================================")
print()

# Função com temporizador para escolha do modo
modo_execucao = ["B"]  # Lista mutável para permitir alteração pela thread

def perguntar_modo():
    escolha = input("Deseja executar de forma visível (V) ou em background/headless (B)? [V/B]: ").strip().upper()
    if escolha in ["V", "B"]:
        modo_execucao[0] = escolha

# Inicia a thread de input
input_thread = threading.Thread(target=perguntar_modo)
input_thread.daemon = True
input_thread.start()

# Aguarda até 15 segundos pela resposta
input_thread.join(timeout=15)

if modo_execucao[0] == "B":
    print()
    print("⏱️ Tempo esgotado ou escolha não feita. Executando em modo background (headless).")
    print()
else:
    print(f"🔄 Executando em modo {'visível' if modo_execucao[0] == 'V' else 'background'}.")

# Configuração para acesso ao Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credenciais_path = os.path.join(os.path.dirname(__file__), 'credenciais.json')
credentials = ServiceAccountCredentials.from_json_keyfile_name(credenciais_path, scope)
client = gspread.authorize(credentials)

spreadsheet_url = "https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing"
sheet_name = "CONTAORDEM"
print(f"Selecionando dados na sheet: {sheet_name}")
sheet = client.open_by_url(spreadsheet_url).worksheet(sheet_name)
data = sheet.get_all_records()

#----------------------------------------------------------------------------------------------

from webdriver_manager.chrome import ChromeDriverManager

chrome_options = Options()

if modo_execucao[0] == "B":
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")

# Inicializar o WebDriver
try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    print("✅ WebDriver inicializado com sucesso.")
except Exception as e:
    print(f"❌ Erro ao iniciar o WebDriver: {e}")
    exit(1)



#(2)#################################################################################

print()
print("==================================================================")
print(f"(2) Iniciando o processo de acesso ao SOMA")
print("==================================================================")
print()

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

# Tentativa de localizar e interagir com o botão SOMA
try:
    botao_soma = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, '285'))
    )
    botao_soma.click()
    print("Botão 'SOMA' clicado com sucesso!")
except TimeoutException:
    print("Erro: Botão 'SOMA' não encontrado ou não ficou disponível a tempo!")
 
# Aguarda o carregamento da página após o clique
try:
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
    )
    print("Página carregada com sucesso após clicar no botão 'SOMA'.")
except TimeoutException:
    print("Erro: Página não carregou após clicar no botão 'SOMA'.")


#(3)##################################################################################

print()
print("==================================================================")
print(f"(3) Iniciando o processo de interração com a leitura dos dados da sheet {sheet_name}")
print("==================================================================")
print()
print("Processando todas as linhas!")
print()

for idx, linha in enumerate(data):

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
            #print(f"Processando a linha {index}: {row}")
                        
            # Escrever "Em processamento" na coluna 'DOC. SOMA' da linha atual
            cabecalho = sheet.row_values(1)  # Obter cabeçalho da planilha
            if 'DOC. SOMA' in cabecalho:
                coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1  # Obter índice 1-based da coluna
                sheet.update_cell(index, coluna_doc_soma, "Em processamento")
                #print(f"Linha {index} validada e marcada 'Em processamento' na coluna 'DOC. SOMA'.")
                                
            else:
                print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
                continue
    else:
        print("=============================================================================")
        print()
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA').")
        print()
        print("=============================================================================")
        
        # Chamar processo de atualização de caixas "SYS_CAIXA".
        atualizar_caixas()
    
    # Imprimir os nomes das colunas disponíveis e o total de linhas filtradas para depuração
    if linhas_filtradas:
        
        #print(f"Colunas disponíveis na planilha: {data[0].keys()}")
        print(f"Total de linhas filtradas (DOC. SOMA = vazio): {len(linhas_filtradas)}")
        
    else:
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA'")
        sys.exit(0)
  # Interrompe a execução do script

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

    #(4)#################################################################################

    print()
    print("==================================================================")
    print(f"(4) Iniciando o processo de interração com a página do SOMA")
    print("==================================================================")
    print()


    if tipo_xl != "Transferência":

              
        # URL inicial
        url_inicial = "https://verbodavida.info/IVV/"

        # Função para redirecionar para a página inicial e validar o carregamento
        def redirecionar_para_inicio(driver, url, timeout=20):
            print()
            print("==================================================================")
            print(f"(3.1) Reinicalizando o processo de interração com a leitura dos dados da sheet CONTAORDEM")
            print("==================================================================")
            print()
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
        
        
        # Localizar e clicar na opção "Entradas/Saídas"
        try:
            elemento = WebDriverWait(driver, 20).until(  # Aguarda até 20 segundos ou até que o elemento esteja clicável
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
            )
            elemento.click()
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
       
    #-------------------------------------------------------------------------------------
           
    if tipo_xl == "Saída":
         
        # Preencher os campos na página "Saída"
        
        botao_radio_saida = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[1]/input')
        time.sleep(3)
        botao_radio_saida.click()
        print("Selecionado o botão para o processo de 'Saída'")
        time.sleep(1)
        
        print(" ")
        print("(4.1) ===================================================")
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
             
        
        print(" ")
        print("(4.2) ===================================================")
        print(f"Iniciando o processo de inserir pagamento para a linha {index}")
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
        print("(4.3) ===================================================")
        print(f"Iniciando o processo de baixa do documento para a linha {index}")
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

            # Aguardar até que o modal SweetAlert2 desapareça antes de clicar no botão Voltar
            try:
                WebDriverWait(driver, 10).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, 'swal2-container'))
                )
                print("Modal SweetAlert desapareceu com sucesso.")
            except TimeoutException:
                print("Aviso: O modal SweetAlert ainda está visível após o tempo de espera.")

            # Clicar no botão 'Voltar'
            try:
                botao_voltar = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]'))
                )
                botao_voltar.click()  # Clicar no botão 'Voltar'
                print("Clicou na opção 'Entradas/Saídas' com sucesso!")
                time.sleep(1)  # Aguarda para garantir que o formulário seja processado antes de continuar
            except TimeoutException:
                print("Erro: Botão 'Entradas/Saídas' não ficou clicável a tempo!")
                time.sleep(2)



        print(" ")
        print("(4.4) ===================================================")
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
        
               
        print(" ")
        print("(4.5) ===================================================")
        print(f"Iniciando o processo de pesquisa da descrção dos inputs do nº documento: {doc_soma}")
        print(" ")

        # Construir o URL para a nova página com base no 'doc_soma'
        novo_url = f"https://verbodavida.info/IVV/?mod=ivv&exec=entradas_saidas_dados&ID={doc_soma}"
        print(f"Redirecionando para: {novo_url}")

        # Navegar para a nova página
        driver.get(novo_url)

        # Esperar o carregamento da nova página
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]'))
            )
            print("Nova página carregada com sucesso!")
        except TimeoutException:
            print("Nova página não carregou no tempo esperado!")

        # Obter os dados do documento
        try:
            dados_elemento = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]')
            if dados_elemento:
                dados_doc = dados_elemento.text.strip()  # Captura o texto do elemento
                print(f"Número do documento extraído: {dados_doc}")
            else:
                print("Elemento 'dados_elemento' não encontrado ou vazio.")
        except Exception as e:
            print(f"Erro ao extrair dados do documento: {e}")


        import datetime

        # Obter o cabeçalho da planilha
        cabecalho = sheet.row_values(1)

        # Verificar se as colunas necessárias existem no cabeçalho
        try:
            coluna_dados_doc = cabecalho.index("DADOS DOC") + 1  # Índice baseado em 1
            coluna_iduser = cabecalho.index("IDUSER") + 1
            coluna_timestamp = cabecalho.index("TIMESTAMP") + 1
        except ValueError as e:
            print(f"Erro ao localizar colunas na planilha: {e}")
            exit(1)

        # Gravar os valores na planilha
        try:
            # Atualizar a coluna DADOS DOC
            sheet.update_cell(index, coluna_dados_doc, dados_doc)
            print(f"Descrição '{dados_doc}' gravada na linha {index}, coluna 'DADOS DOC'.")

            # Atualizar a coluna IDUSER
            sheet.update_cell(index, coluna_iduser, "USERJOB")
            print(f"IDUSER 'USERJOB' gravado na linha {index}, coluna 'IDUSER'.")

            # Atualizar a coluna TIMESTAMP
            timestamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            sheet.update_cell(index, coluna_timestamp, timestamp)
            print(f"TIMESTAMP '{timestamp}' gravado na linha {index}, coluna 'TIMESTAMP'.")

        except Exception as e:
            print(f"Erro ao atualizar as colunas: {e}")

        
        print(" ")
        print(f"Processamento finalizado para a linha {index}")
        print("(4.6) ===================================================")
        print(" ")

        # Redirecionar para a página inicial após processar o documento
        try:
            redirecionar_para_inicio(driver, url_inicial)
        except TimeoutException as e:
            print(f"Erro ao redirecionar para a página inicial: {e}")
            #return  # Finalizar o processamento se não for possível voltar à página inicial
    
    #-------------------------------------------------------------------------------------

    if tipo_xl == "Entrada":
                    
        # Preencher os campos na página "Entrada"
            
        botao_radio_entrada = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[2]/input')
        time.sleep(3)
        botao_radio_entrada.click()
        print("Selecionado o botão de rádio 'Entrada'")
        time.sleep(1)
        
        #5.1#################################################################################
                
        print(" ")
        print("(5.1) ===================================================")
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
                print(f"A forma de pagamento é diferente de DINHEIRO. pular a lógia de processamento do caixa entrada.")

        
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
        
        #5.1#################################################################################
        
        print(" ")
        print("(5.1) ===================================================")
        print(f"Iniciando o processo de baixa do documento para a linha {index}")
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
                botao_OK_baixa = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
                )
                botao_OK_baixa.click()
                print("Botão 'OK Baixa' clicado com sucesso!")
                time.sleep(2)
            except TimeoutException:
                print("Erro: Botão 'OK Baixa' não encontrado ou não ficou clicável a tempo!")

            # Tentar aguardar o desaparecimento do SweetAlert, mas continuar mesmo se ele não sumir
            try:
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, 'swal2-container'))
                )
                print("Modal SweetAlert desapareceu com sucesso.")
            except TimeoutException:
                print("Aviso: O modal SweetAlert ainda está visível após o tempo de espera.")

            # Tentar clicar no botão 'Voltar' — forçar com JavaScript se necessário
            try:
                botao_voltar = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]'))
                )
                try:
                    botao_voltar.click()  # Tentar normalmente
                    print("Clicou na opção 'Entradas/Saídas' com sucesso!")
                except ElementClickInterceptedException:
                    print("Clique no botão 'Voltar' foi interceptado. Forçando com JavaScript...")
                    driver.execute_script("arguments[0].click();", botao_voltar)
                    print("Clique forçado no botão 'Voltar' realizado com sucesso.")
                time.sleep(1)
            except TimeoutException:
                print("Erro: Botão 'Voltar' não ficou clicável a tempo.")
                time.sleep(2)


                
            print( )
           
        else:
            print("Forma de pagamento é diferente de 'TRANSFERÊNCIA BANCÁRIA'. Pulando o processo de baixa de pagamento.")
    

        #5.2 #################################################################################
            
        print("(5.2) ===================================================")
        print(f"Iniciando o processo pesquisa do nº soma do documento")
        print(" ")

        # Clicar botão voltar a sessão "Entradas/Saídas"
           

        try:
            # Esperar até que o botão esteja clicável
            botao_voltar = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]'))
            )
            botao_voltar.click()  # Clicar no botão 'Voltar'
            print("Clicou na opção 'Entradas/Saídas' com sucesso!")
            time.sleep(1)  # Aguarda para garantir que o formulário seja processado antes de continuar
        except TimeoutException:
            print("Erro: Botão 'Entradas/Saídas' não ficou clicável a tempo!")
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
        
        #5.3 #################################################################################

        # Processo principal
        print(" ")
        print("(5.3) ===================================================")
        print(f"Iniciando o processo pesquisa DESCRIÇÃO do documento: {doc_soma}")
        print(" ")

        # Construir o URL para a nova página com base no 'doc_soma'
        novo_url = f"https://verbodavida.info/IVV/?mod=ivv&exec=entradas_saidas_dados&ID={doc_soma}"
        print(f"Redirecionando para: {novo_url}")

        # Navegar para a nova página
        driver.get(novo_url)

        # Esperar o carregamento da nova página
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]'))
            )
            print("Nova página carregada com sucesso!")
        except TimeoutException:
            print("Nova página não carregou no tempo esperado!")

        # Obter os dados do documento
        try:
            dados_elemento = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]')
            if dados_elemento:
                dados_doc = dados_elemento.text.strip()  # Captura o texto do elemento
                print(f"Número do documento extraído: {dados_doc}")
            else:
                print("Elemento 'dados_elemento' não encontrado ou vazio.")
        except Exception as e:
            print(f"Erro ao extrair dados do documento: {e}")

 
        import datetime

        # Obter o cabeçalho da planilha
        cabecalho = sheet.row_values(1)

        # Verificar se as colunas necessárias existem no cabeçalho
        try:
            coluna_dados_doc = cabecalho.index("DADOS DOC") + 1  # Índice baseado em 1
            coluna_iduser = cabecalho.index("IDUSER") + 1
            coluna_timestamp = cabecalho.index("TIMESTAMP") + 1
        except ValueError as e:
            print(f"Erro ao localizar colunas na planilha: {e}")
            exit(1)

        # Gravar os valores na planilha
        try:
            # Atualizar a coluna DADOS DOC
            sheet.update_cell(index, coluna_dados_doc, dados_doc)
            print(f"Descrição '{dados_doc}' gravada na linha {index}, coluna 'DADOS DOC'.")

            # Atualizar a coluna IDUSER
            sheet.update_cell(index, coluna_iduser, "USERJOB")
            print(f"IDUSER 'USERJOB' gravado na linha {index}, coluna 'IDUSER'.")

            # Atualizar a coluna TIMESTAMP
            timestamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            sheet.update_cell(index, coluna_timestamp, timestamp)
            print(f"TIMESTAMP '{timestamp}' gravado na linha {index}, coluna 'TIMESTAMP'.")

        except Exception as e:
            print(f"Erro ao atualizar as colunas: {e}")


        print(" ")
        print(f"Processamento finalizado para a linha {index}")
        print("(5.4) ===================================================")
        print(" ")

        # Redirecionar para a página inicial após processar o documento
        try:
            redirecionar_para_inicio(driver, url_inicial)
        except TimeoutException as e:
            print(f"Erro ao redirecionar para a página inicial: {e}")
            #return  # Finalizar o processamento se não for possível voltar à página inicial

    #-------------------------------------------------------------------------------------      
               
    if tipo_xl == "Transferência":
        print(" ")
        print("🔥🔥🔥🔥🔥🔥🔥🔥 O tipo do documento é Transferência. Continuando o processamento...")
        print(" ")


        # URL inicial
        url_inicial = "https://verbodavida.info/IVV/"

        # Função para redirecionar para a página inicial e validar o carregamento
        def redirecionar_para_inicio(driver, url, timeout=20):
            print()
            print("==================================================================")
            print(f"(3.1) Reinicalizando o processo de interração com a leitura dos dados da sheet CONTAORDEM")
            print("==================================================================")
            print()
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

        print(" ")
        print(f"Processamento finalizado para a linha {index}")
        print("(7.0) ===================================================")
        print(" ")

        # Redirecionar para a página inicial após processar o documento
        try:
            redirecionar_para_inicio(driver, url_inicial)
        except TimeoutException as e:
            print(f"Erro ao redirecionar para a página inicial: {e}")
            #return  # Finalizar o processamento se não for possível voltar à página inicial

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
        root = tk.Tk()
        root.title("Fim do Processamento")
        root.geometry("350x150")
        root.eval('tk::PlaceWindow . center')  # Centra a janela
        label = tk.Label(root, text="O processamento foi finalizado com sucesso!", wraplength=300)
        label.pack(pady=20)

        btn = tk.Button(root, text="Fechar", command=root.destroy)
        btn.pack()

        root.mainloop()  # Essencial para manter a janela aberta até clique
    except Exception as e:
        print(f"Erro ao exibir popup: {e}")

# Chamar a função para exibir o popup
mostrar_popup()

# Encerrar o programa
sys.exit(0)

