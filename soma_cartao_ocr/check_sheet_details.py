#!/usr/bin/env python3
"""Script para verificar os dados exatos da sheet"""

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

    # Ler sheet
    response = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{worksheet}'!A1:M30"
    ).execute()

    sheet_data = response.get("values", [])

    print("\n" + "="*120)
    print("DETALHES DA SHEET CARTAO")
    print("="*120)

    print(f"\nTotal de linhas: {len(sheet_data)}\n")

    # Mostrar cada linha com indice
    for idx, linha in enumerate(sheet_data, 1):
        vazio = all(str(v).strip() == "" for v in linha) if linha else True
        status = "[VAZIA]" if vazio else "[DADOS]"
        print(f"Linha {idx:2d} {status}: {linha[:5]}...")

    print("\n" + "="*120)
    print("\nVERIFICAO ESPECIFICA:")
    print("  - Linha 1 deve ter CABECALHO")
    print("  - Linhas 2-4 podem estar vazias ou com dados anteriores")
    print("  - Linha 5 deve estar VAZIA")
    print("  - Linha 6 deve ter primeiro movimento")
    print("  - Linha 7+ devem ter proximos movimentos")

    if len(sheet_data) > 4:
        linha_5_vazia = all(str(v).strip() == "" for v in sheet_data[4])
        print(f"\nStatus Linha 5: {'OK - VAZIA' if linha_5_vazia else 'ERRO - NAO ESTA VAZIA'}")

    if len(sheet_data) > 5:
        linha_6_vazia = all(str(v).strip() == "" for v in sheet_data[5])
        print(f"Status Linha 6: {'VAZIA' if linha_6_vazia else 'TEM DADOS'}")
        if not linha_6_vazia:
            print(f"  Primeira coluna (Data Mov): '{sheet_data[5][0]}'")

    print("\n" + "="*120 + "\n")

if __name__ == "__main__":
    main()
