###################################################################################
# selenium_scripts/processar_transferencia.py (AJUSTADO PARA EVITAR INTERCEPTAÇÃO)
###################################################################################

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
import time

# Função utilitária original
def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
    except:
        return None

# Função para realizar scroll até o elemento
def scroll_to_element(driver, elemento):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    time.sleep(0.5)

# Função de redirecionamento original
def redirecionar_para_inicio(driver, url, timeout=20):
    print()
    print("==================================================================")
    print(f"(3.1) Reinicializando o processo de interação com a leitura dos dados")
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

def processar_transferencia(driver, linha, index, sheet):
    
    # Extração de variáveis da folha
    try:
        descricao_xl = str(linha.get('DESCRIÇÃO SOMA', ''))
        data_xl = str(linha.get('DATA MOV.', ''))
        valor_xl = str(linha.get('IMPORTÂNCIA', ''))
        caixa_xl = str(linha.get('CAIXA', ''))
        caixa_saida_xl = str(linha.get('CAIXA SAIDA', ''))
        tipo_xl = str(linha.get('TIPO', ''))
    except KeyError as e:
        print(f"Erro: A coluna '{e.args[0]}' não foi encontrada nesta linha. Dados ignorados.")
        return

    print(" ")
    print("🔥🔥🔥🔥🔥🔥🔥🔥 O tipo do documento é Transferência. Continuando o processamento...")
    print(" ")

    # URL inicial
    url_inicial = "https://verbodavida.info/IVV/"

    try:
        redirecionar_para_inicio(driver, url_inicial)
    except TimeoutException as e:
        print(f"Erro ao redirecionar para a página inicial: {e}")
        return
                    
    # Localizar e clicar na opção "Transferências Caixas"
    try:
        caixas_bancos = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div/div[4]/div/div/span'))
        )
        time.sleep(2)
        caixas_bancos.click()
        print("Clicou na opção 'caixas bancos' com sucesso!")
    except TimeoutException:
        print("Erro: Elemento 'caixas bancos' não encontrado ou não ficou clicável a tempo!")
        
    # Após carregar a página de Transferências, clicar no botão "Nova Transferência"
    try:
        nova_transferencia = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/a'))
        )
        time.sleep(2)
        nova_transferencia.click()
        print("Clicou no botão 'Nova Entrada/Saída' com sucesso!")
    except TimeoutException:
        print("Erro: Botão 'Nova Entrada/Saída' não encontrado ou não ficou clicável a tempo!")

                 
    # Selecionar o campo "Caixa Saída" - (Transferência)
    caixa_saida = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[2]/div')
    if caixa_saida:
        time.sleep(1)
        caixa_saida.click()  # Clica no campo para abrir o dropdown
        
        # Localizar o campo de busca no dropdown
        selecao_caixa_saida = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
        )
        
        # Preencher o valor e confirmar com ENTER
        selecao_caixa_saida.clear()
        selecao_caixa_saida.send_keys(caixa_saida_xl)
        WebDriverWait(driver, 5).until(
            EC.text_to_be_present_in_element_value(
                (By.XPATH, '/html/body/span/span/span[1]/input'), caixa_saida_xl
            )
        )
        
        selecao_caixa_saida.send_keys(Keys.ENTER)
        print(f"CAIXA DIÁRIO preenchido com sucesso: {caixa_saida_xl}")
    else:
        print("Erro: Campo para selecionar CAIXA DIÁRIO não encontrado.")
        
    time.sleep(1)

    # Selecionar o campo "Valor"
    valor_transferencia = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[4]/div/input')
    if valor_transferencia:
        valor_transferencia.send_keys(str(valor_xl))
        time.sleep(1)
        print(f"Valor preenchido com sucesso: {valor_xl}")
    else:
        print("Campo de valor não encontrado!")

    time.sleep(1)

    # Selecionar o campo "Caixa Entrada" - (Transferência)
    caixa_entrada = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[5]/div')
    if caixa_entrada:
        time.sleep(1)
        caixa_entrada.click()
        
        # Localizar o campo de busca no dropdown
        selecao_caixa_entrada = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/span/span/span[1]/input'))
        )
        
        # Preencher o valor e confirmar com ENTER
        selecao_caixa_entrada.clear()
        selecao_caixa_entrada.send_keys(caixa_xl)
        WebDriverWait(driver, 5).until(
            EC.text_to_be_present_in_element_value(
                (By.XPATH, '/html/body/span/span/span[1]/input'), caixa_xl
            )
        )
        
        selecao_caixa_entrada.send_keys(Keys.ENTER)
        print(f"CAIXA DIÁRIO preenchido com sucesso: {caixa_xl}")
    else:
        print("Erro: Campo para selecionar CAIXA DIÁRIO não encontrado.")

    time.sleep(1)
    
    # Selecionar o campo "Data"
    data_trensferencia = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[8]/div/div/input')
    if data_trensferencia:
        data_trensferencia.clear()
        data_trensferencia.send_keys(data_xl)
        time.sleep(0.5)
        data_trensferencia.send_keys(Keys.ESCAPE) # FECHAR O CALENDÁRIO para não interceptar o clique no Salvar
        print(f"Data de transferência preenchida com sucesso: {data_xl}")
    else:
        print("Campo de data de transferencia não encontrado!")

    time.sleep(1)

    # Selecionar o campo "Descrição"
    descricao_campo = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[9]/div/textarea')
    if descricao_campo:
        descricao_campo.send_keys(descricao_xl)
        print(f"Descrição preenchida com sucesso: {descricao_xl}")
    else:
        print("Campo de descrição não encontrado!")

    # Salvar transferência - (Saída)
    try:
        botao_salvar_transferencia = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[10]/div/button[1]'))
        )
        scroll_to_element(driver, botao_salvar_transferencia)
        
        try:
            botao_salvar_transferencia.click()
        except ElementClickInterceptedException:
            # Caso o calendário ainda esteja na frente, forçar clique via JavaScript
            print("Clique interceptado, tentando via JavaScript...")
            driver.execute_script("arguments[0].click();", botao_salvar_transferencia)
            
        print("Salvar pagamento salvo com sucesso!")
    except TimeoutException:
        print("Erro: Botão 'Salvar pagamento' não encontrado ou não ficou clicável a tempo!")

    
    # Clicar no botão "OK A transferência foi salva" - (Saída)
    try:
        botao_OK_pagamento = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/button[1]'))
        )
        time.sleep(2) # Pequena pausa para a animação do modal
        botao_OK_pagamento.click()
        print("Botão OK:'A transferência foi salva' ")
    except TimeoutException:
        print("Erro: Botão 'OK transferencia' não encontrado ou não ficou clicável a tempo!")   

    # Clicar botão voltar a sessão "Entradas/Saídas"
    botao_voltar = tentar_encontrar_elemento(driver, By.XPATH, '/html/body/div[2]/div/div[2]/div/div/form/div[10]/div/button[2]')
    if botao_voltar:
        scroll_to_element(driver, botao_voltar)
        try:
            botao_voltar.click()
        except:
            driver.execute_script("arguments[0].click();", botao_voltar)
        print("Clicou na opção 'Entradas/Saídas' com sucesso!")
        time.sleep(1)
    else:
        print("Elemento 'Entradas/Saídas' não encontrado!")
    
    # Atualizar o valor na planilha após processar a linha
    cabecalho = sheet.row_values(1)
    if 'DOC. SOMA' in cabecalho:
        coluna_doc_soma = cabecalho.index('DOC. SOMA') + 1
        texto_fixo = "Transferido"
        sheet.update_cell(index, coluna_doc_soma, texto_fixo)
        print(f"Texto fixo '{texto_fixo}' gravado na linha {index}, coluna 'DOC. SOMA'.")
    else:
        print("Erro: Coluna 'DOC. SOMA' não encontrada no cabeçalho da planilha!")
        
    print(f"\n✅ Transferência finalizada com sucesso para a linha {index}!")