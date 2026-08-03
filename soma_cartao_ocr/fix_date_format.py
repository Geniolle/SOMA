#!/usr/bin/env python3
"""Corrige formatação de datas na sheet CARTÃO."""

import re
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
import yaml

config_path = Path(__file__).parent / "config.yaml"
with open(config_path, encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

creds_path = Path(cfg["google"]["service_account_file"])
if not creds_path.is_absolute():
    creds_path = (Path(__file__).parent / creds_path).resolve()

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
]

credentials = service_account.Credentials.from_service_account_file(
    str(creds_path),
    scopes=scopes,
)

sheets_service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
spreadsheet_id = cfg["google_sheets"]["spreadsheet_id"]

print("📅 Verificando formatação de datas na sheet CARTÃO...\n")

# Ler dados da sheet CARTÃO
response = sheets_service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range="'CARTÃO'!A1:Z1000",
).execute()

values = response.get("values", [])
if not values:
    print("❌ Sheet CARTÃO está vazia!")
    exit(1)

header = values[0]
header_lower = [h.lower() for h in header]

# Encontrar colunas de data
d_mov_col = next((i for i, h in enumerate(header_lower) if "data mov" in h), None)
d_val_col = next((i for i, h in enumerate(header_lower) if "data valor" in h), None)

print(f"Coluna 'Data Mov.': índice {d_mov_col} ({header[d_mov_col] if d_mov_col is not None else 'NÃO ENCONTRADA'})")
print(f"Coluna 'Data Valor': índice {d_val_col} ({header[d_val_col] if d_val_col is not None else 'NÃO ENCONTRADA'})\n")

if d_mov_col is None or d_val_col is None:
    print("❌ Colunas de data não encontradas!")
    exit(1)

# Padrões de data
iso_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}$")  # YYYY-MM-DD
ddmm_pattern = re.compile(r"^\d{2}/\d{2}$")  # DD/MM
ddmmyyyy_pattern = re.compile(r"^\d{2}/\d{2}/\d{4}$")  # DD/MM/YYYY

dates_to_fix = []

print("Verificando datas...\n")
print(f"{'Linha':<6} {'Data Mov.':<20} {'Data Valor':<20} {'Status':<15}")
print("-" * 65)

for row_idx, row in enumerate(values[1:], 2):
    d_mov = str(row[d_mov_col]).strip() if len(row) > d_mov_col else ""
    d_val = str(row[d_val_col]).strip() if len(row) > d_val_col else ""

    status = "✅ OK"

    # Verificar Data Mov
    if d_mov and iso_pattern.match(d_mov):
        print(f"{row_idx:<6} {d_mov:<20} {d_val:<20} {'⚠️  ISO FORMAT':<15}")
        # Converter YYYY-MM-DD para DD/MM/YYYY
        year, month, day = d_mov.split("-")
        new_d_mov = f"{day}/{month}/{year}"
        dates_to_fix.append((row_idx, "Data Mov.", d_mov_col, d_mov, new_d_mov))
        status = "⚠️  CORRIGIR"
    elif d_mov and ddmm_pattern.match(d_mov):
        status = "✅ DD/MM"
    elif d_mov and ddmmyyyy_pattern.match(d_mov):
        status = "✅ DD/MM/YYYY"

    if status != "✅ OK":
        print(f"{row_idx:<6} {d_mov:<20} {d_val:<20} {status:<15}")

print()

if dates_to_fix:
    print(f"\n⚠️  {len(dates_to_fix)} data(s) em formato ISO encontrada(s):\n")

    for row_idx, col_name, col_idx, old_date, new_date in dates_to_fix:
        print(f"Linha {row_idx}: {col_name}")
        print(f"  '{old_date}' → '{new_date}'")

    # Perguntar se quer corrigir
    if input("\nDeseja corrigir? (s/n): ").lower() == "s":
        print("\nCorrigindo datas...\n")

        for row_idx, col_name, col_idx, old_date, new_date in dates_to_fix:
            col_letter = chr(65 + col_idx)

            # Atualizar na sheet
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"'CARTÃO'!{col_letter}{row_idx}",
                valueInputOption="USER_ENTERED",
                body={"values": [[new_date]]},
            ).execute()

            print(f"✅ Linha {row_idx}: {col_letter}{row_idx} = '{new_date}'")

        print("\n✅ Datas corrigidas com sucesso!")
else:
    print("✅ Todas as datas estão no formato correto (DD/MM ou DD/MM/YYYY)!")
