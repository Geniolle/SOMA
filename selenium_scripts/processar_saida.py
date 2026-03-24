###################################################################################
# selenium_scripts/processar_saida.py
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
        raise

def processar_saida(driver, linha, index, sheet):
    
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

    # BLOCO TRY-EXCEPT GERAL PARA MARCAR 'ERROR' NA COLUNA STATUS EM CASO DE FALHA
    try:
        # URL inicial e Navegação Inicial
        url_inicial = "https://verbodavida.info/IVV/"
        
        try:
            redirecionar_para_inicio(driver, url_inicial)
        except TimeoutException as e:
            print(f"Erro ao redirecionar para a página inicial: {e}")
            raise e

        # Localizar e clicar na opção "Entradas/Saídas"
        try:
            elemento = WebDriverWait(driver, 20).until(  
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
            )
            elemento.click()
            print("Clicou na opção 'Entradas/Saídas' com sucesso!")
        except TimeoutException:
            print("Erro: Elemento 'Entradas/Saídas' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)

        # Após carregar a página de Entradas/Saídas, clicar no botão "Nova Entrada/Saída"
        try:
            nova_entrada_saida = WebDriverWait(driver, 20).until(  
                EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "btn btn-primary") and contains(., "Nova Entrada/Saída")]'))
            )
            nova_entrada_saida.click()
            print("Clicou no botão 'Nova Entrada/Saída' com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Nova Entrada/Saída' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)


        # =================================================================================
        # INÍCIO DO CÓDIGO DA "SAÍDA"
        # =================================================================================
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
            plano_de_conta = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span/span[1]')
            if not plano_de_conta: 
                plano_de_conta = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span')
                
            if plano_de_conta:
                plano_de_conta.click()  
                selecao_plano_de_conta = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, '//input[@class="select2-search__field"]'))
                )
                selecao_plano_de_conta.clear()  
                selecao_plano_de_conta.send_keys(plano_de_conta_xl)  
                WebDriverWait(driver, 5).until(  
                    EC.text_to_be_present_in_element_value(
                        (By.XPATH, '//input[@class="select2-search__field"]'), plano_de_conta_xl
                    )
                )
                selecao_plano_de_conta.send_keys(Keys.ENTER)  
                print(f"Plano de conta preenchido com sucesso: {plano_de_conta_xl}")
            else:
                print("Erro: Campo para selecionar plano de conta não encontrado.")
                time.sleep(2)

            # Preencher o campo "Centro de custo"
            centro_de_custo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[8]/div/span/span[1]/span')
            if not centro_de_custo: 
                centro_de_custo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span')
                
            if centro_de_custo:
                centro_de_custo.click()  
                selecao_centro_de_custo = WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, '//input[@class="select2-search__field"]'))
                )
                selecao_centro_de_custo.clear()  
                selecao_centro_de_custo.send_keys(centro_custo_xl)  
                WebDriverWait(driver, 2).until(  
                    EC.text_to_be_present_in_element_value(
                        (By.XPATH, '//input[@class="select2-search__field"]'), centro_custo_xl
                    )
                )
                selecao_centro_de_custo.send_keys(Keys.ENTER)  
                print(f"Centro de custo preenchido com sucesso: {centro_custo_xl}")
            else:
                print("Erro: Campo para selecionar Centro de custo não encontrado.")
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
                time.sleep(1)  
                valor_campo.send_keys(Keys.ENTER)  
                print(f"Valor preenchido com sucesso: {valor_xl}")
            else:
                print("Campo de valor não encontrado!")
                time.sleep(2)
            
            # --- PREENCHIMENTO DA DESCRIÇÃO ---
            descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/input')
            if not descricao_campo: 
                descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/input')
                
            if descricao_campo:
                descricao_campo.send_keys(descricao_xl)
                print(f"Descrição preenchida com sucesso: {descricao_xl}")
                time.sleep(1)
            else:
                print("Campo de descrição não encontrado!")
                time.sleep(2)
                
            # Selecionar o campo "Obs"
            observacao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[21]/div/textarea')
            if observacao_campo:
                observacao_campo.click()
                observacao_campo.send_keys(descricao_xl)
                print(f"Observação preenchida com sucesso: {descricao_xl}")
            else:
                print("Campo de observação não encontrado!")
                time.sleep(2)
                 
            # --- NOVO BLOCO ADICIONADO: SALVAR O FORMULÁRIO PRINCIPAL ---
            botao_salvar_principal = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[30]/div/button[1]')
            if botao_salvar_principal:
                botao_salvar_principal.click()
                print("Formulário salvo com sucesso!")
                time.sleep(5)  # Aguarda o sistema processar a gravação
            else:
                print("Botão de salvar o formulário não encontrado!")
                time.sleep(2)
            # ------------------------------------------------------------
            
            print(" ")
            print("(4.2) ===================================================")
            print(f"Iniciando o processo de inserir pagamento para a linha {index}")
            print(" ")

            # Realizar pagamento
            botao_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/ul/li/div/div[2]/button[1]')
            if botao_pagamento:
                botao_pagamento.click()  
                print("Realizar pagamento salvo com sucesso!")
                time.sleep(5)  
            else:
                print("Botão de Realizar pagamento não encontrado!")
                time.sleep(2)
            
            # Inserir pagamento
            botao_inserir_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[29]/div/div[2]/a')
            if botao_inserir_pagamento:
                botao_inserir_pagamento.click()  
                print("Inserir pagamento salvo com sucesso!")
                time.sleep(2)  
            else:
                print("Botão de Inserir pagamento não encontrado!")
                time.sleep(2)
            
            # Selecionar o campo "Data" no Modal de Pagamento
            data_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input')
            if data_pagamento:
                data_pagamento.clear()  
                data_pagamento.send_keys(data_xl)  
                print(f"Data início preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data inicial não encontrado!")
            time.sleep(1)
            
            # Selecionar o campo "Forma de pagamento" - (Saída)
            forma_de_pagamento1 = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select')
            if forma_de_pagamento1:
                try:
                    select_forma_pagamento1 = Select(forma_de_pagamento1)
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
                    select_caixa_pagamento = Select(caixa_pagamento)
                    select_caixa_pagamento.select_by_visible_text(caixa_xl)
                    print(f"Caixa para pagamento selecionada com sucesso: {caixa_xl}")
                except Exception as e:
                    print(f"Erro ao selecionar a Caixa para pagamento: {e}")
            else:
                print("Campo 'Caixa para pagamento' não encontrado!")
            time.sleep(2)

            # Verificar se a forma de pagamento é "TRANSFERÊNCIA BANCÁRIA"
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
                botao_Salvar_pagamento = WebDriverWait(driver, 20).until(  
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[5]/div/div/form/div[3]/button'))
                )
                botao_Salvar_pagamento.click()  
                print("Botão 'Salvar Pagamento' clicado com sucesso!")
            except TimeoutException:
                print("Erro: Botão 'Salvar Pagamento' não encontrado ou não ficou clicável a tempo!")   
            time.sleep(2)

            # Clicar no botão "OK pagamento"
            try:
                botao_OK_pagamento = WebDriverWait(driver, 20).until(  
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
                )
                time.sleep(5)
                botao_OK_pagamento.click()  
                print("Botão 'OK Pagamento' clicado com sucesso!")
            except TimeoutException:
                print("Erro: Botão 'OK Pagamento' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)

            
            print(" ")
            print("(4.3) ===================================================")
            print(f"Iniciando o processo de baixa do documento para a linha {index}")
            print(" ")

            if forma_de_pagamento_xl == "TRANSFERÊNCIA BANCÁRIA":
                print("A forma de pagamento é TRANSFERÊNCIA BANCÁRIA. Continuando o processamento...")
                    
                # Inserir baixa
                botao_inserir_baixa = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button')
                if botao_inserir_baixa:
                    botao_inserir_baixa.click()  
                    print("Inserir Baixa salvo com sucesso!")
                    time.sleep(2)  
                else:
                    print("Botão de Inserir Baixa não encontrado!")
                        
                    
                # Selecionar o campo "Data"
                data_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input')
                if data_pagamento:
                    data_pagamento.clear()  
                    data_pagamento.send_keys(data_xl)  
                    print(f"Data início preenchida com sucesso: {data_xl}")
                    time.sleep(2)
                else:
                    print("Campo de data inicial não encontrado!")
                    

                # Clicar no pop-up "Baixa"
                observacao_campo_baixa = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[3]')
                if observacao_campo_baixa:
                    observacao_campo_baixa.click()
                    print(f"Click na janela pop-up com sucesso")
                    time.sleep(2)
                else:
                    print("Campo de descrição não encontrado!")
                
                
                # Salvar baixa
                try:
                    botao_salvar_baixa = WebDriverWait(driver, 20).until(  
                        EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[3]/button'))
                    )
                    botao_salvar_baixa.click()  
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
                    botao_voltar.click()  
                    print("Clicou na opção 'Voltar' com sucesso!")
                    time.sleep(1)  
                except TimeoutException:
                    print("Erro: Botão 'Voltar' não ficou clicável a tempo.")
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
                data_intervaloA.clear()  
                data_intervaloA.send_keys(data_xl)  
                print(f"Data início preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data inicial não encontrado!")
            time.sleep(1)

            data_intervaloB = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[3]/div/input')
            if data_intervaloB:
                data_intervaloB.clear()  
                data_intervaloB.send_keys(data_xl)  
                print(f"Data fim preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data final não encontrado!")
            time.sleep(1)

            # Após carregar a página de Entradas/Saídas, clicar no botão "Pesquisar"
            try:
                pesquisar_documento = WebDriverWait(driver, 20).until(  
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/button'))
                )
                pesquisar_documento.click()  
                print("Botão 'Pesquisar' clicado com sucesso!")
            except TimeoutException:
                print("Erro: Botão 'Pesquisar' não encontrado ou não ficou clicável a tempo!")
                time.sleep(2)
        
            # Obter o número do documento gerado após salvar
            texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/div/div/table/tbody/tr/td[3]')
            
            # Correção segura para garantir que a coluna é encontrada antes de atualizar
            cabecalho = sheet.row_values(1)
            coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1 if 'DOC. SOMA' in cabecalho else 1

            if texto_elemento:
                doc_soma = texto_elemento.text.strip()  
                print(f"Número do documento extraído: {doc_soma}")

                # Atualizar o valor na planilha após processar a linha
                if 'DOC. SOMA' in cabecalho:
                    sheet.update_cell(index, coluna_doc_soma, doc_soma)  
                    print(f"Texto '{doc_soma}' gravado na linha {index}, coluna 'DOC. SOMA'.")
                else:
                    print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
            else:
                print(f"Não foi possível extrair o número do documento para a linha {index}.")
                sheet.update_cell(index, coluna_doc_soma, "Em erro")
                print(f"Processamento concluído para a linha {index}.")
                time.sleep(2)
                doc_soma = "000000"
            
                
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
                    dados_doc = dados_elemento.text.strip()  
                    print(f"Número do documento extraído: {dados_doc}")
                else:
                    print("Elemento 'dados_elemento' não encontrado ou vazio.")
            except Exception as e:
                print(f"Erro ao extrair dados do documento: {e}")

            # Obter o cabeçalho da planilha
            cabecalho = sheet.row_values(1)

            # Verificar se as colunas necessárias existem no cabeçalho
            try:
                coluna_dados_doc = cabecalho.index("DADOS DOC") + 1  
                coluna_iduser = cabecalho.index("IDUSER") + 1
                coluna_timestamp = cabecalho.index("TIMESTAMP") + 1
            except ValueError as e:
                print(f"Erro ao localizar colunas na planilha: {e}")
                return

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

    # CAPTURA DE QUALQUER ERRO PARA GRAVAR "ERROR" NA COLUNA STATUS
    except Exception as erro_geral:
        print(f"\n❌ Erro crítico capturado durante o processamento da linha {index}: {erro_geral}")
        try:
            cabecalho = sheet.row_values(1)
            if "STATUS" in cabecalho:
                coluna_status = cabecalho.index("STATUS") + 1
                sheet.update_cell(index, coluna_status, "ERROR")
                print(f"⚠️ 'ERROR' gravado com sucesso na coluna STATUS da linha {index}.")
            else:
                print("⚠️ Aviso: Coluna 'STATUS' não encontrada na folha. Impossível gravar o erro.")
        except Exception as erro_planilha:
            print(f"⚠️ Falha ao tentar gravar 'ERROR' na planilha: {erro_planilha}")
        
        # Levanta a exceção novamente para não dar como um processo bem-sucedido
        raise erro_geral