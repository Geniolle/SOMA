#!/usr/bin/env python3
"""Script para validar a sheet CARTAO contra resultado.json"""

import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
import yaml

def load_config(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}

def load_credentials(cfg: dict):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds_path = Path(cfg.get("google", {}).get("service_account_file"))
    if not creds_path.is_absolute():
        creds_path = (Path(__file__).parent / creds_path).resolve()
    return service_account.Credentials.from_service_account_file(str(creds_path), scopes=scopes)

def main():
    cfg = load_config(Path("config.yaml"))
    creds = load_credentials(cfg)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)

    spreadsheet_id = cfg.get("google_sheets", {}).get("spreadsheet_id")
    worksheet = cfg.get("google_sheets", {}).get("worksheet", "CARTAO")

    # Ler resultado.json
    with open("output/resultado.json", encoding="utf-8") as f:
        resultado = json.load(f)

    # Ler sheet
    response = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{worksheet}'!A1:M30"
    ).execute()

    sheet_data = response.get("values", [])

    print("\n" + "="*100)
    print("VALIDACAO DA SHEET CARTAO")
    print("="*100 + "\n")

    print(f"Sheet: {worksheet}")
    print(f"Linhas na sheet: {len(sheet_data)}")
    print(f"Movimentos em resultado.json: {len(resultado['movimentos'])}\n")

    # Verificar cabeçalho
    if sheet_data:
        print(f"CABECALHO (linha 1):")
        print(f"{sheet_data[0]}\n")

    # Verificar linha 5 (deve estar vazia)
    print(f"\nVERIFICACAO LINHA 5 (DEVE ESTAR VAZIA):")
    print("-"*100)
    if len(sheet_data) > 4:
        linha_5 = sheet_data[4]  # indice 4 = linha 5 (1-based)
        print(f"Dados na sheet linha 5: {linha_5}")

        # Verificar se está vazia
        vazio = all(str(v).strip() == "" for v in linha_5)
        status = "OK - VAZIA (CORRETO)" if vazio else "ERRO - NAO ESTA VAZIA"
        print(f"Status: {status}\n")
    else:
        print(f"Sheet tem menos de 5 linhas\n")

    # Verificar primeiros movimentos (linhas 6-10)
    print(f"\nVERIFICACAO LINHAS 6-10 (PRIMEIROS MOVIMENTOS):")
    print("-"*100)

    if len(sheet_data) > 5:
        for i in range(5, min(10, len(sheet_data))):
            linha = i + 1  # 1-based line number
            dados_sheet = sheet_data[i]

            # Encontrar movimento correspondente em resultado.json
            mov_json = next((m for m in resultado['movimentos'] if m['line'] == linha), None)

            print(f"\nLinha {linha}:")
            if mov_json:
                print(f"  esperado em JSON  : linha {mov_json['line']}")
                print(f"  data_movimento    : '{mov_json['data_movimento']}' (JSON)")
                if len(dados_sheet) > 0:
                    print(f"  data_movimento    : '{dados_sheet[0]}' (SHEET)")
                print(f"  descricao         : '{mov_json['descricao']}' (JSON)")
                if len(dados_sheet) > 2:
                    print(f"  descricao         : '{dados_sheet[2]}' (SHEET)")
                print(f"  debito_eur        : '{mov_json['debito_eur']}' (JSON)")
                if len(dados_sheet) > 7:
                    print(f"  debito_eur        : '{dados_sheet[7]}' (SHEET)")
                print(f"  Status JSON       : {mov_json['status']}")

                # Validacoes
                problemas = []

                # Verificar se contem termos de erro como "Taxa", "Debita", "Original"
                for col_idx, valor in enumerate(dados_sheet):
                    val_str = str(valor).strip()
                    if val_str in ["Taxa", "Debita", "Original", "Cambio", "EUR ( )"]:
                        problemas.append(f"Coluna {col_idx}: contem palavra de cabecalho '{val_str}'")

                if problemas:
                    print(f"  PROBLEMAS DETECTADOS:")
                    for p in problemas:
                        print(f"      - {p}")
            else:
                print(f"  ATENCAO: Movimento linha {linha} nao encontrado em resultado.json")
                print(f"  Dados na sheet: {dados_sheet}")

    print("\n" + "="*100 + "\n")

if __name__ == "__main__":
    main()
