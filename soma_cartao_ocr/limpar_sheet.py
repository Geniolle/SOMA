#!/usr/bin/env python3
"""Script para limpar a sheet CARTAO removendo todos os dados exceto o cabecalho"""

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
    print("\nLIMPEZA DA SHEET CARTAO")
    print("="*80)
    print("\nEste script deletara TODOS os dados da sheet CARTAO, mantendo apenas o cabecalho.\n")

    cfg = load_config(Path("config.yaml"))
    creds = load_credentials(cfg)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)

    spreadsheet_id = cfg.get("google_sheets", {}).get("spreadsheet_id")
    worksheet = cfg.get("google_sheets", {}).get("worksheet", "CARTAO")

    # Ler sheet atual
    response = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{worksheet}'!A1:Z5000"
    ).execute()

    sheet_data = response.get("values", [])
    print(f"Linhas na sheet: {len(sheet_data)}")

    if not sheet_data:
        print("Sheet esta vazia. Nada a fazer.")
        return

    # Manter apenas o cabecalho (primeira linha)
    print(f"\nCabecalho: {sheet_data[0]}")
    print(f"\nA DELETAR: {len(sheet_data) - 1} linhas de dados")

    if len(sheet_data) > 1:
        print("\nTentando deletar linhas 2 em diante...")

        # Obter sheet ID
        meta = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = next((s["properties"]["sheetId"] for s in meta.get("sheets", []) if s["properties"]["title"] == worksheet), None)

        if sheet_id is not None:
            # Deletar todas as linhas exceto a primeira
            requests = [
                {
                    "deleteRange": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 1,
                            "endRowIndex": len(sheet_data),
                        },
                        "shiftDimension": "ROWS"
                    }
                }
            ]

            sheets.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()

            print("Linhas deletadas com sucesso!\n")

            # Verificar resultado
            response = sheets.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{worksheet}'!A1:Z10"
            ).execute()

            final_data = response.get("values", [])
            print(f"Status final da sheet:")
            print(f"  Linhas: {len(final_data)}")
            if final_data:
                print(f"  Cabecalho: {final_data[0]}")

            print("\n" + "="*80)
            print("SHEET LIMPA COM SUCESSO!")
            print("="*80)
            print("\nProximos passos:")
            print("1. Execute: python main.py")
            print("2. Execute: python validate_sheet.py")
            print("="*80 + "\n")
        else:
            print("ERRO: Nao foi possivel encontrar o ID da sheet")

if __name__ == "__main__":
    main()
