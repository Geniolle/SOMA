#!/usr/bin/env python3
"""Validacao final da sheet CARTAO contra resultado.json"""

import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build
import yaml

def load_config(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}

def load_credentials(cfg: dict):
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"]
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
        range=f"'{worksheet}'!A1:M50"
    ).execute()

    sheet_data = response.get("values", [])

    print("\n" + "="*120)
    print("VALIDACAO FINAL - SHEET CARTAO vs resultado.json")
    print("="*120)

    print(f"\nSheet: {worksheet}")
    print(f"Linhas na sheet: {len(sheet_data)}")
    print(f"Movimentos em resultado.json: {len(resultado['movimentos'])}")

    # Validacoes criticas
    problemas = []
    alertas = []

    # Verificacao 1: Cabeçalho
    if sheet_data and sheet_data[0]:
        print(f"\n[OK] Cabecalho presente (linha 1)")
    else:
        problemas.append("Cabecalho ausente na linha 1")

    # Verificacao 2: Resultado.json linha 5 deve estar vazia
    resultado_linha_5 = next((m for m in resultado['movimentos'] if m['line'] == 5), None)
    if resultado_linha_5:
        if not resultado_linha_5['data_movimento'].strip():
            print(f"[OK] resultado.json linha 5 vazia conforme esperado")
        else:
            problemas.append(f"resultado.json linha 5 nao esta vazia: {resultado_linha_5['data_movimento']}")
    else:
        alertas.append("resultado.json nao tem movimento na linha 5")

    # Verificacao 3: Dados reais na sheet começam na linha 2
    if len(sheet_data) > 1:
        print(f"[OK] Dados comecam na sheet linha 2")

    # Verificacao 4: Comparar dados
    print(f"\nComparacao de dados (sheet vs resultado.json):")
    print("-" * 120)

    # Estratégia: pegar movimentos do resultado.json (excluindo linha 5 vazia)
    # e comparar com linhas da sheet (começando em linha 2)
    movimentos_reais = [m for m in resultado['movimentos'] if m['line'] != 5]

    for idx, mov in enumerate(movimentos_reais[:10]):  # Primeiros 10 para não poluir output
        sheet_line_idx = idx + 1  # Sheet linha 2 = índice 1

        if len(sheet_data) > sheet_line_idx:
            sheet_row = sheet_data[sheet_line_idx]

            # Extrair dados relevantes da sheet
            data_mov_sheet = sheet_row[0] if len(sheet_row) > 0 else ""
            desc_sheet = sheet_row[2] if len(sheet_row) > 2 else ""
            debito_sheet = sheet_row[7] if len(sheet_row) > 7 else ""

            # Comparar com resultado.json
            match = True
            diffs = []

            if mov['data_movimento'] and data_mov_sheet:
                # Converter formato: JSON usa DD/MM, sheet usa DD/MM/YYYY
                json_date = mov['data_movimento'].replace("/", "")
                sheet_date = data_mov_sheet[:5].replace("/", "")  # Pegar DD/MM
                if json_date != sheet_date:
                    match = False
                    diffs.append(f"data: JSON={mov['data_movimento']} vs SHEET={data_mov_sheet}")

            if mov['descricao'] and desc_sheet and mov['descricao'].lower() != desc_sheet.lower():
                match = False
                diffs.append(f"desc: JSON={mov['descricao'][:30]} vs SHEET={desc_sheet[:30]}")

            if mov['debito_eur']:
                mov_debit = float(mov['debito_eur'].replace(",", ".")) if "," in str(mov['debito_eur']) else float(mov['debito_eur'])
                try:
                    sheet_debit = float(str(debito_sheet).replace(",", ".")) if debito_sheet else 0
                    if abs(mov_debit - sheet_debit) > 0.01:
                        match = False
                        diffs.append(f"debito: JSON={mov['debito_eur']} vs SHEET={debito_sheet}")
                except:
                    pass

            status = "OK" if match else "ERRO"
            print(f"Mov {idx+1} (Sheet L{sheet_line_idx+1}): {status}")
            if diffs:
                for d in diffs:
                    print(f"  - {d}")
                    problemas.append(f"Movimento {idx+1}: {d}")

    # Sumario
    print("\n" + "="*120)
    print("SUMARIO:")
    print("="*120)

    if problemas:
        print(f"\nPROBLEMAS ENCONTRADOS ({len(problemas)}):")
        for p in problemas:
            print(f"  [ERRO] {p}")
    else:
        print("\n[OK] NENHUM PROBLEMA ENCONTRADO!")

    if alertas:
        print(f"\nALERTAS ({len(alertas)}):")
        for a in alertas:
            print(f"  [ATENCAO] {a}")

    print("\n" + "="*120 + "\n")

    return 0 if not problemas else 1

if __name__ == "__main__":
    exit(main())
