import gspread
import sys
import time
import os
import smtplib
import tkinter as tk
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
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def mostrar_popup(mensagem):
    """
    Exibe um popup com a mensagem fornecida.
    """
    try:
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal
        messagebox.showinfo("Status do Processamento", mensagem)
        root.quit()
    except Exception as e:
        print(f"Erro ao exibir popup: {e}")

def atualizar_caixas():
    print("Atualizando caixas...")

    #(Bloco <Inicio do Processo (6)>)#################################################################################

    print()
    print("==================================================================")
    print(f"(6) Iniciando o processo 'Caixas/bancos'")
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
        return (1)

    # Autorização do Google Sheets
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credenciais_path, scope)
        client = gspread.authorize(credentials)
    except Exception as e:
        print(f"Erro ao autorizar credenciais do Google Sheets: {e}")
        return (1)

    # URL do Google Sheets e nome da sheet
    spreadsheet_url = "https://docs.google.com/spreadsheets/d/1poVWJGSBb13_2S1YKEzvFmkB9Ru0ZVzfQ0OEcMkfOZw/edit?usp=sharing"
    sheet_name = "GERENCIAR CAIXAS"

    try:
    
        sheet = client.open_by_url(spreadsheet_url).worksheet(sheet_name)
        print(f"Selecionando dados na sheet: {sheet_name}\n")
        data = sheet.get_all_records()
    except Exception as e:
        print(f"Erro ao acessar a planilha ou a sheet: {e}")
        return (1)

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
        return (1)

    # Inicializar o WebDriver
    try:
        driver = webdriver.Chrome(service=Service(webdriver_path), options=chrome_options)
        #print("WebDriver inicializado com sucesso.")
    except Exception as e:
        print("Erro ao iniciar o WebDriver:")
        print(f"- Verifique o caminho do ChromeDriver: {webdriver_path}")
        print(f"- Certifique-se de que a versão do ChromeDriver é compatível com o navegador.")
        print(f"Erro detalhado: {e}")
        return (1)


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

      #(Bloco <Fim do Processo (6)>)#################################################################################


    
    #(Bloco <Inicio do Processo (2)>)#################################################################################

    # URL do site para automação
    url = 'https://verbodavida.info/apps/index.php'

    #print(f"Acessando o site: {url}")
    driver.get(url)

    # Iniciando o processo de login
    #print("Iniciando o processo de login...")

    # Login no site
    usuario = tentar_encontrar_elemento(driver, By.NAME, 'email')
    if usuario:
        usuario.send_keys('familialopesemportugal@gmail.com')
        #print("Campo 'email' preenchido com sucesso.")
    else:
        print("Erro: Campo 'email' não encontrado!")

    senha = tentar_encontrar_elemento(driver, By.NAME, 'senha')
    if senha:
        senha.send_keys('P@1internet')
        #print("Campo 'senha' preenchido com sucesso.")
    else:
        print("Erro: Campo 'senha' não encontrado!")

    botao_login = tentar_encontrar_elemento(driver, By.NAME, 'submit')
    if botao_login:
        botao_login.click()
        #print("Botão 'login' clicado com sucesso.")
        time.sleep(2)  # Tempo para garantir o carregamento da página seguinte
    else:
        print("Erro: Botão 'login' não encontrado!")

    # Tentativa de localizar e interagir com o botão SOMA
    try:
        botao_soma = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, '285'))
        )
        botao_soma.click()
        #print("Botão 'SOMA' clicado com sucesso!")
    except TimeoutException:
        print("Erro: Botão 'SOMA' não encontrado ou não ficou disponível a tempo!")
        driver.save_screenshot("erro_soma.png")

    for idx, linha in enumerate(data):
    
        # Aguarda o carregamento da página após o clique
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
            )
            #print("Página carregada com sucesso após clicar no botão 'SOMA'.")
        except TimeoutException:
            #print("Erro: Página não carregou após clicar no botão 'SOMA'.")
            driver.save_screenshot("erro_pagina_soma.png")
    
    #(Bloco <Inicio do Processo (2)>)#################################################################################



    #(Bloco <Inicio do Processo (3)>)#################################################################################
    

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
                    #print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
                    continue
        else:
            print("Não há linhas para processar (todas têm valor em 'DOC. SOMA').")

    #(Bloco <Inicio do Processo (3)>)#################################################################################


    
        #(Bloco <Inicio do Processo (6.1)>)#################################################################################

        print()
        print("==================================================================")
        print(f"(6.1) Iniciando o processo de atualização 'Caixas/bancos'")
        print("==================================================================")
        print()
        
        
        try:
            botao_caixas_bancos = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div/div[1]/div'))
            )
            botao_caixas_bancos.click()
            print("Clicou na opção 'Caixas/bancos' com sucesso!")
        except TimeoutException:
            print("Erro: Elemento 'Caixas/bancos' não encontrado ou não ficou clicável a tempo!")

        
        print(" ")
        #(Bloco <Final do Processo (6.1)>)#################################################################################

        #(Bloco <Inicio do Processo (6.2)>)#################################################################################

        print(" ")
        print("(6.2) ===================================================")
        print(f"Processando colunas específicas")
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
            print()
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
                    print()
                else:
                    print("Erro: Coluna 'D. LIVRARIA' não encontrada no cabeçalho da planilha.")
            except ValueError as e:
                print(f"Erro ao encontrar a coluna: {e}")
            except Exception as e:
                print(f"Erro ao atualizar a sheet: {e}")
        else:
            print("Erro: Não foi possível extrair o valor para 'D. LIVRARIA'.")

        return ()  # Interrompe a execução do script

    # Imprimir os nomes das colunas disponíveis e o total de linhas filtradas para depuração
    if linhas_filtradas:
        
        #print(f"Colunas disponíveis na planilha: {data[0].keys()}")
        print(f"Total de linhas filtradas (DOC. SOMA = vazio): ({len(linhas_filtradas)})")
        
    else:
        print("Não há linhas para processar (todas têm valor em 'DOC. SOMA'")
        return ()  # Interrompe a execução do script


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
            
    #(Bloco <Inicio do Processo (6.2)>)#################################################################################


    # Finalizar o WebDriver
    try:
        driver.quit()
        print("WebDriver finalizado com sucesso.")
    except Exception as e:
        print(f"Erro ao finalizar o WebDriver: {e}")

    def mostrar_popup():
        try:
            root = tk.Tk()
            root.withdraw()  # Ocultar a janela principal
            messagebox.showinfo("Processamento Concluído", "O processamento foi finalizado com sucesso.")
            root.update_idletasks()
            time.sleep(1)  # Adiciona um atraso para garantir que o popup seja exibido
            root.quit()
        except Exception as e:
            print(f"Erro ao exibir popup: {e}")

    # Chamar a função para exibir o popup
    mostrar_popup()

    # Encerrar o programa
    print("Processamento Concluído: O processamento foi finalizado com sucesso.")
    sys.exit(0)

