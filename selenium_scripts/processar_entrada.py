###################################################################################
# selenium_scripts/processar_entrada.py
###################################################################################

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time
import datetime

# Função utilitária base
def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, valor)))
    except:
        return None

# Função para fazer scroll seguro para o centro do ecrã
def scroll_to_element(driver, elemento):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    time.sleep(0.4) 

# Função Master para lidar com os dropdowns chatos do Select2
def preencher_select2(driver, elemento_container, texto_pesquisa):
    scroll_to_element(driver, elemento_container)
    try:
        elemento_container.click()
    except Exception:
        driver.execute_script("arguments[0].click();", elemento_container)
        time.sleep(0.3)
        driver.execute_script("var ev = new MouseEvent('mousedown', {bubbles: true}); arguments[0].dispatchEvent(ev);", elemento_container)
    
    selecao = WebDriverWait(driver, 5).until(
        EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
    )
    selecao.clear()
    selecao.send_keys(texto_pesquisa)
    WebDriverWait(driver, 5).until(
        EC.text_to_be_present_in_element_value((By.XPATH, '/html/body/span/span/span[1]/input'), texto_pesquisa)
    )
    selecao.send_keys(Keys.ENTER)
    time.sleep(0.5)

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
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
        )
        print("Página inicial carregada com sucesso!")
    except TimeoutException:
        print("Erro: Não foi possível carregar a página inicial.")
        raise

def processar_entrada(driver, linha, index, sheet):
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

    try:
        url_inicial = "https://verbodavida.info/IVV/"
        try:
            redirecionar_para_inicio(driver, url_inicial)
        except TimeoutException as e:
            print(f"Erro ao redirecionar para a página inicial: {e}")
            raise e

        try:
            elemento = WebDriverWait(driver, 20).until(  
                EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"Entradas/saídas")]'))
            )
            elemento.click()
            print("Clicou na opção 'Entradas/Saídas' com sucesso!")
        except TimeoutException:
            print("Erro: Elemento 'Entradas/Saídas' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)

        try:
            nova_entrada_saida = WebDriverWait(driver, 20).until(  
                EC.element_to_be_clickable((By.XPATH, '//a[contains(@class, "btn btn-primary") and contains(., "Nova Entrada/Saída")]'))
            )
            nova_entrada_saida.click()
            print("Clicou no botão 'Nova Entrada/Saída' com sucesso!")
        except TimeoutException:
            print("Erro: Botão 'Nova Entrada/Saída' não encontrado ou não ficou clicável a tempo!")
            time.sleep(2)

        if tipo_xl == "Entrada":
            botao_radio_entrada = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[2]/input')
            time.sleep(3)
            try:
                botao_radio_entrada.click()
            except ElementClickInterceptedException:
                 driver.execute_script("arguments[0].click();", botao_radio_entrada)
            print("Selecionado o botão de rádio 'Entrada'")
            time.sleep(1)
            
            print(" ")
            print("(5.1) ===================================================")
            print(f"Iniciando o processo de input de dados para a linha {index} e o tipo {tipo_xl}")
            print(" ")

            plano_de_conta = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span/span[1]')
            if not plano_de_conta: 
                plano_de_conta = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/span/span[1]/span/span[1]')
            if plano_de_conta:
                preencher_select2(driver, plano_de_conta, plano_de_conta_xl)
                print(f"Plano de conta preenchido com sucesso: {plano_de_conta_xl}")
            else:
                print("Erro: Campo para selecionar plano de conta não encontrado.")

            centro_de_custo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[8]/div/span/span[1]/span')
            if not centro_de_custo: 
                centro_de_custo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span')
            if centro_de_custo:
                preencher_select2(driver, centro_de_custo, centro_custo_xl)
                print(f"Centro de custo preenchido com sucesso: {centro_custo_xl}")
            else:
                print("Erro: Campo para selecionar Centro de custo não encontrado.")
            
            data_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[11]/div/div/input')
            if data_campo:
                scroll_to_element(driver, data_campo)
                data_campo.send_keys(data_xl)
                print(f"Data preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data não encontrado!")
            
            valor_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[12]/div/input')
            if valor_campo:
                scroll_to_element(driver, valor_campo)
                valor_campo.send_keys(str(valor_xl))
                time.sleep(1)  
                valor_campo.send_keys(Keys.ENTER)  
                print(f"Valor preenchido com sucesso: {valor_xl}")
            else:
                print("Campo de valor não encontrado!")
                    
            forma_de_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[16]/div/span/span[1]/span/span[1]')
            if not forma_de_pagamento:
                forma_de_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[16]/div/span')
            if forma_de_pagamento:
                preencher_select2(driver, forma_de_pagamento, forma_de_pagamento_xl)
                print(f"Forma de pagamento preenchida com sucesso: {forma_de_pagamento_xl}")
            else:
                print("Erro: Campo para selecionar Forma de pagamento não encontrado.")
            
            
            # --- NOVA LÓGICA DE CAIXA: USANDO A TAG <SELECT> ORIGINAL ---
            print(f"A forma de pagamento é {forma_de_pagamento_xl}. Selecionar o campo Caixa.")
            
            caixa_select_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[20]/div/select')
            if caixa_select_elemento:
                try:
                    # Injeta a seleção diretamente no HTML, ignorando máscaras visuais (Select2)
                    script_js = f"""
                        var select = arguments[0];
                        var textoProcurado = '{caixa_xl}';
                        for (var i = 0; i < select.options.length; i++) {{
                            if (select.options[i].text.includes(textoProcurado)) {{
                                select.selectedIndex = i;
                                // Dispara o evento de 'change' para que a página saiba que mudámos o valor
                                select.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                break;
                            }}
                        }}
                    """
                    driver.execute_script(script_js, caixa_select_elemento)
                    print(f"Caixa '{caixa_xl}' selecionado diretamente no elemento <select> com sucesso!")
                    time.sleep(1)
                except Exception as e:
                    print(f"Erro ao forçar a seleção do Caixa via JS: {e}")
            else:
                print("Erro: Campo <select> do Caixa não encontrado com o XPath fornecido.")
            
            
            # --- PREENCHIMENTO DA DESCRIÇÃO COM OCULTAÇÃO DA CAIXA DE SUGESTÕES ---
            descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/input')
            if not descricao_campo:
                descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/input')
                
            if descricao_campo:
                scroll_to_element(driver, descricao_campo)
                descricao_campo.clear()
                descricao_campo.send_keys(descricao_xl)
                print(f"Descrição preenchida com sucesso: {descricao_xl}")
                time.sleep(0.5)
                descricao_campo.send_keys(Keys.ESCAPE) 
                
                # Forçamos o painel de sugestões a ficar invisível
                js_hide_sugestao = """
                var xpath = "/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/div/div";
                var result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
                if (result.singleNodeValue) {
                    result.singleNodeValue.style.display = 'none';
                }
                """
                driver.execute_script(js_hide_sugestao)
                time.sleep(0.5)
            else:
                print("Campo de descrição não encontrado!")

            # Selecionar o campo "Obs"
            observacao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[21]/div/textarea')
            if observacao_campo:
                scroll_to_element(driver, observacao_campo)
                try:
                    observacao_campo.click()
                except:
                    driver.execute_script("arguments[0].click();", observacao_campo)
                observacao_campo.clear()
                observacao_campo.send_keys(descricao_xl)
                print(f"Observação preenchida com sucesso: {descricao_xl}")
            else:
                print("Campo de observação não encontrado!")
                 
            # Salvar o formulário
            botao_salvar_principal = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[30]/div/button[1]')
            if botao_salvar_principal:
                scroll_to_element(driver, botao_salvar_principal)
                driver.execute_script("arguments[0].click();", botao_salvar_principal)
                print("Formulário salvo com sucesso!")
                time.sleep(5)
            else:
                print("Botão de salvar o formulário não encontrado!")
            
            #5.1#################################################################################
            
            print(" ")
            print("(5.1) ===================================================")
            print(f"Iniciando o processo de baixa do documento para a linha {index}")
            print(" ")

            if forma_de_pagamento_xl == "TRANSFERÊNCIA BANCÁRIA":
                print("A forma de pagamento é TRANSFERÊNCIA BANCÁRIA. iniciar processo de baixa de pagamento...")
                print(" ")

                botao_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/ul/li/div/div[2]/button[1]')
                if botao_pagamento:
                    scroll_to_element(driver, botao_pagamento)
                    try:
                        botao_pagamento.click()
                    except:
                        driver.execute_script("arguments[0].click();", botao_pagamento)
                    print("Realizar pagamento salvo com sucesso!")
                    time.sleep(2)
                else:
                    print("Botão de Realizar pagamento não encontrado!")
                    
                botao_inserir_baixa = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button')
                if botao_inserir_baixa:
                    scroll_to_element(driver, botao_inserir_baixa)
                    try:
                        botao_inserir_baixa.click()
                    except:
                         driver.execute_script("arguments[0].click();", botao_inserir_baixa)
                    print("Inserir Baixa salvo com sucesso!")
                    time.sleep(2)
                else:
                    print("Botão de Inserir Baixa não encontrado!")
                    
                data_pagamento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input')
                if data_pagamento:
                    data_pagamento.clear()
                    data_pagamento.send_keys(data_xl)
                    print(f"Data início preenchida com sucesso: {data_xl}")
                else:
                    print("Campo de data inicial não encontrado!")
                time.sleep(1)

                observacao_campo_baixa = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[2]')
                if observacao_campo_baixa:
                    try:
                        observacao_campo_baixa.click()
                    except:
                        driver.execute_script("arguments[0].click();", observacao_campo_baixa)
                    print(f"Click na janela pop-up com sucesso")
                    time.sleep(2)
                else:
                    print("Campo do pop-up não encontrado!")

                try:
                    botao_salvar_baixa = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[4]/div/div/form/div[3]/button'))
                    )
                    driver.execute_script("arguments[0].click();", botao_salvar_baixa)
                    print("Salvar Baixa salvo com sucesso!")
                    time.sleep(2)
                except TimeoutException:
                    print("Erro: Botão 'Salvar Baixa' não encontrado ou não ficou clicável a tempo!")
                    
                try:
                    botao_OK_baixa = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div/button[1]'))
                    )
                    driver.execute_script("arguments[0].click();", botao_OK_baixa)
                    print("Botão 'OK Baixa' clicado com sucesso!")
                    time.sleep(2)
                except TimeoutException:
                    print("Erro: Botão 'OK Baixa' não encontrado ou não ficou clicável a tempo!")

                try:
                    WebDriverWait(driver, 10).until_not(
                        EC.presence_of_element_located((By.CLASS_NAME, 'swal2-container'))
                    )
                    print("Modal SweetAlert desapareceu com sucesso.")
                except TimeoutException:
                    print("Aviso: O modal SweetAlert ainda está visível após o tempo de espera.")

                try:
                    botao_voltar = WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]'))
                    )
                    driver.execute_script("arguments[0].click();", botao_voltar)
                    print("Clicou na opção 'Voltar' com sucesso!")
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

            pesquisar_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/input')
            if pesquisar_campo:
                scroll_to_element(driver, pesquisar_campo)
                pesquisar_campo.send_keys(descricao_xl)
                print(f"Campo pesquisar a descrição preenchida com sucesso: {descricao_xl}")
            else:
                print("Campo de pesquisar a descrição não encontrado!")
            time.sleep(1)
                
            botao_radio_periodo = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/div/input')
            driver.execute_script("arguments[0].click();", botao_radio_periodo)
            print("Selecionado o botão de rádio 'Periodo'")
            time.sleep(1)
            
            botao_radio_data_pagamento = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[8]/div/div[2]/input')
            driver.execute_script("arguments[0].click();", botao_radio_data_pagamento)
            print("Selecionado o botão de rádio 'Data de Pagamento'")
            time.sleep(1)

            data_intervaloA = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[1]/div/input')
            if data_intervaloA:
                data_intervaloA.clear()
                data_intervaloA.send_keys(data_xl)
                time.sleep(0.5)
                data_intervaloA.send_keys(Keys.ESCAPE)
                print(f"Data início preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data inicial não encontrado!")
            time.sleep(1)

            data_intervaloB = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[3]/div/input')
            if data_intervaloB:
                data_intervaloB.clear()
                data_intervaloB.send_keys(data_xl)
                time.sleep(0.5)
                data_intervaloB.send_keys(Keys.ESCAPE) 
                print(f"Data fim preenchida com sucesso: {data_xl}")
            else:
                print("Campo de data final não encontrado!")
            time.sleep(1)

            try:
                pesquisar_documento = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/button'))
                )
                driver.execute_script("arguments[0].click();", pesquisar_documento)
                print("Botão 'Pesquisar' clicado via JS com sucesso!")
            except TimeoutException:
                print("Erro: Botão 'Pesquisar' não encontrado ou não ficou clicável a tempo!")
                time.sleep(2)
        
            texto_elemento = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[4]/div/div/div/div/table/tbody/tr/td[3]')
            
            cabecalho = sheet.row_values(1)
            coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1 if 'DOC. SOMA' in cabecalho else 1

            if texto_elemento:
                doc_soma = texto_elemento.text.strip()
                print(f"Número do documento extraído: {doc_soma}")

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
            
            #5.3 #################################################################################

            print(" ")
            print("(5.3) ===================================================")
            print(f"Iniciando o processo pesquisa DESCRIÇÃO do documento: {doc_soma}")
            print(" ")

            novo_url = f"https://verbodavida.info/IVV/?mod=ivv&exec=entradas_saidas_dados&ID={doc_soma}"
            print(f"Redirecionando para: {novo_url}")

            driver.get(novo_url)

            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]'))
                )
                print("Nova página carregada com sucesso!")
            except TimeoutException:
                print("Nova página não carregou no tempo esperado!")

            try:
                dados_elemento = driver.find_element(By.XPATH, '/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]')
                if dados_elemento:
                    dados_doc = dados_elemento.text.strip()
                    print(f"Número do documento extraído: {dados_doc}")
                else:
                    print("Elemento 'dados_elemento' não encontrado ou vazio.")
            except Exception as e:
                print(f"Erro ao extrair dados do documento: {e}")

     
            cabecalho = sheet.row_values(1)

            try:
                coluna_dados_doc = cabecalho.index("DADOS DOC") + 1
                coluna_iduser = cabecalho.index("IDUSER") + 1
                coluna_timestamp = cabecalho.index("TIMESTAMP") + 1
            except ValueError as e:
                print(f"Erro ao localizar colunas na planilha: {e}")
                return

            try:
                sheet.update_cell(index, coluna_dados_doc, dados_doc)
                print(f"Descrição '{dados_doc}' gravada na linha {index}, coluna 'DADOS DOC'.")

                sheet.update_cell(index, coluna_iduser, "USERJOB")
                print(f"IDUSER 'USERJOB' gravado na linha {index}, coluna 'IDUSER'.")

                timestamp = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                sheet.update_cell(index, coluna_timestamp, timestamp)
                print(f"TIMESTAMP '{timestamp}' gravado na linha {index}, coluna 'TIMESTAMP'.")

            except Exception as e:
                print(f"Erro ao atualizar as colunas: {e}")


            print(" ")
            print(f"Processamento finalizado para a linha {index}")
            print("(5.4) ===================================================")
            print(" ")

            try:
                redirecionar_para_inicio(driver, url_inicial)
            except TimeoutException as e:
                print(f"Erro ao redirecionar para a página inicial: {e}")

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
        
        raise erro_geral