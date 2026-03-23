###################################################################################
# selenium_scripts/processar_entrada.py (LÓGICA 100% ORIGINAL REPLICADA)
###################################################################################

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time
import datetime

# Função utilitária original
def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
    except:
        return None

# Função de redirecionamento original
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
        # driver.save_screenshot("erro_pagina_inicial.png") # Desativado apenas para não gerar lixo
        raise

def processar_entrada(driver, linha, index, sheet):
    
    # Extração de variáveis da folha
    try:
        plano_de_conta_xl = str(linha.get('PLANO DE CONTA', ''))
        centro_custo_xl = str(linha.get('CENTRO DE CUSTO', ''))
        descricao_xl = str(linha.get('DESCRIÇÃO SOMA', ''))
        data_xl = str(linha.get('DATA MOV.', ''))
        valor_xl = str(linha.get('IMPORTÂNCIA', ''))
        forma_de_pagamento_xl = str(linha.get('FORMA DE PAGAMENTO', ''))
        caixa_xl = str(linha.get('CAIXA', ''))
        tipo_xl = str(linha.get('TIPO', ''))
    except KeyError as e:
        print(f"Erro: A coluna '{e.args[0]}' não foi encontrada nesta linha. Dados ignorados.")
        return

    # URL inicial
    url_inicial = "https://verbodavida.info/IVV/"
    
    # Lógica de redirecionamento que estava no inicio do loop (4)
    try:
        redirecionar_para_inicio(driver, url_inicial)
    except TimeoutException as e:
        print(f"Erro ao redirecionar para a página inicial: {e}")
        return

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