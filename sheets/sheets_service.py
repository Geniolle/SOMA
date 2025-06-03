###################################################################################
# sheets/sheets_service.py
###################################################################################

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from config.config import CREDENCIAIS_PATH, SPREADSHEET_URL, SHEET_NAME

def obter_sheet():
    """
    Estabelece conexão com o Google Sheets e retorna o objeto da sheet.
    """
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENCIAIS_PATH, scope)
        client = gspread.authorize(creds)

        spreadsheet = client.open_by_url(SPREADSHEET_URL)
        sheet = spreadsheet.worksheet(SHEET_NAME)
        print(f"✅ Worksheet '{SHEET_NAME}' carregada com sucesso.")
        return sheet

    except Exception as e:
        print(f"❌ Erro ao conectar com o Google Sheets: {e}")
        raise

def obter_todos_os_registros(sheet):
    """
    Retorna todas as linhas da folha como lista de dicionários.
    """
    try:
        registros = sheet.get_all_records()
        print(f"✅ {len(registros)} linhas carregadas da folha '{SHEET_NAME}'.")
        return registros
    except Exception as e:
        print(f"❌ Erro ao obter os registros da folha: {e}")
        raise

def atualizar_linha(sheet, indice_linha, dados):
    """
    Atualiza os valores de uma linha específica na folha.
    O índice começa em 2 (linha 1 é cabeçalho).
    """
    try:
        for coluna, valor in dados.items():
            cell = sheet.find(coluna)
            sheet.update_cell(indice_linha, cell.col, valor)
        print(f"✅ Linha {indice_linha} atualizada com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao atualizar linha {indice_linha}: {e}")
        raise
