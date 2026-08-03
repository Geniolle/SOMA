#!/usr/bin/env python3
"""Corrige datas em formato ISO na sheet CARTÃO."""

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

print("📅 Corrigindo datas em formato ISO na sheet CARTÃO...\n")

# Ler dados
response = sheets_service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range="'CARTÃO'!A1:Z1000",
).execute()

values = response.get("values", [])
header = values[0]
header_lower = [h.lower() for h in header]

d_mov_col = next((i for i, h in enumerate(header_lower) if "data mov" in h), None)
d_val_col = next((i for i, h in enumerate(header_lower) if "data valor" in h), None)

iso_pattern = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")

fixes = []

print("Procurando datas em formato ISO...\n")

for row_idx, row in enumerate(values[1:], 2):
    d_mov = str(row[d_mov_col]).strip() if len(row) > d_mov_col else ""

    if d_mov and iso_pattern.match(d_mov):
        match = iso_pattern.match(d_mov)
        year, month, day = match.groups()
        new_d_mov = f"{day}/{month}/{year}"

        print(f"Linha {row_idx}: '{d_mov}' → '{new_d_mov}'")

        # Corrigir
        col_letter = chr(65 + d_mov_col)
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'CARTÃO'!{col_letter}{row_idx}",
            valueInputOption="USER_ENTERED",
            body={"values": [[new_d_mov]]},
        ).execute()

        fixes.append((row_idx, d_mov, new_d_mov))
        print(f"  ✅ Corrigida\n")

if fixes:
    print(f"✅ {len(fixes)} data(s) corrigida(s) com sucesso!")
else:
    print("✅ Nenhuma data em formato ISO encontrada!")
