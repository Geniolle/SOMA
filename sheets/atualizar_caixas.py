###################################################################################
# sheets/atualizar_caixas.py
###################################################################################

import gspread
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from config.config import CREDENCIAIS_PATH, SPREADSHEET_URL

def tentar_encontrar_elemento(driver, by, valor, timeout=20):
    try:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, valor)))
    except:
        return None

def atualizar_caixas_soma(driver):
    print("\n==================================================================")
    print("(6) Iniciando o processo de atualização 'Caixas/bancos'")
    print("==================================================================")

    # 1. Conectar ao Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENCIAIS_PATH, scope)
        client = gspread.authorize(credentials)
        sheet = client.open_by_url(SPREADSHEET_URL).worksheet("GERENCIAR CAIXAS")
        print("✔️ Conectado à folha 'GERENCIAR CAIXAS' com sucesso.")
        cabecalho = sheet.row_values(1)
    except Exception as e:
        print(f"❌ Erro ao conectar ao Google Sheets (GERENCIAR CAIXAS): {e}")
        return

    # 2. Navegar para a página inicial e clicar em Caixas/Bancos
    try:
        driver.get("https://verbodavida.info/IVV/")
        time.sleep(2)
        
        botao_caixas_bancos = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div[1]/div[2]/div/div/div/div[1]/div'))
        )
        botao_caixas_bancos.click()
        print("✔️ Clicou na opção 'Caixas/bancos' com sucesso!")
        time.sleep(3)
    except Exception as e:
        print(f"❌ Erro ao navegar para 'Caixas/bancos': {e}")
        return

    # 3. Extrair e Guardar Valores
    print("\n(6.2) Processando colunas específicas...")
    
    # Dicionário com os 5 caixas (O VERBO CAFÉ no div[4] e o VERBO SHOP no div[5])
    valores_caixas = {
        "CAIXA DIÁRIO": '/html/body/div[2]/div/div[3]/div/div[1]/div/div/div[2]/div[1]/span[2]',
        "CAIXA BANCO": '/html/body/div[2]/div/div[3]/div/div[2]/div/div/div[2]/div[1]/span[2]',
        "D. CRIANÇAS": '/html/body/div[2]/div/div[3]/div/div[3]/div/div/div[2]/div[1]/span[2]',
        "VERBO CAFÉ": '/html/body/div[2]/div/div[3]/div/div[4]/div/div/div[2]/div[1]/span[2]',
        "VERBO SHOP": '/html/body/div[2]/div/div[3]/div/div[5]/div/div/div[2]/div[1]/span[2]'
    }

    for nome_caixa, xpath in valores_caixas.items():
        elemento = tentar_encontrar_elemento(driver, By.XPATH, xpath)
        if elemento:
            valor_extraido = elemento.text.strip()
            print(f" - Valor extraído para '{nome_caixa}': {valor_extraido}")
            
            if nome_caixa in cabecalho:
                coluna_idx = cabecalho.index(nome_caixa) + 1
                try:
                    sheet.update_cell(2, coluna_idx, valor_extraido) # Atualiza sempre a Linha 2
                    print(f"   ✔️ Atualizado no Google Sheets (Linha 2, Coluna '{nome_caixa}')")
                except Exception as e:
                    print(f"   ❌ Erro ao escrever no Sheets: {e}")
            else:
                print(f"   ❌ Erro: Coluna '{nome_caixa}' não encontrada no cabeçalho.")
        else:
            print(f" ❌ Erro: Não foi possível extrair o valor para '{nome_caixa}'.")

    print("\n✅ Processo de atualização de Caixas/bancos finalizado!")