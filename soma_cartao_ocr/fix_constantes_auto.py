#!/usr/bin/env python3
"""Corrige automaticamente o mapeamento na sheet CONSTANTES."""

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
    "https://www.googleapis.com/auth/drive.readonly",
]

credentials = service_account.Credentials.from_service_account_file(
    str(creds_path),
    scopes=scopes,
)

sheets_service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
spreadsheet_id = cfg["google_sheets"]["spreadsheet_id"]

print("📊 Corrigindo sheet CONSTANTES automaticamente...\n")

# Ler dados
response = sheets_service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id,
    range="'CONSTANTES'!A1:Z5000",
).execute()

values = response.get("values", [])
header = values[0]
header_lower = [h.lower() for h in header]

forma_col = next((i for i, h in enumerate(header_lower) if "forma de pagamento" in h), None)
tipo_col = next((i for i, h in enumerate(header_lower) if h == "tipo"), None)

print(f"Corrigindo {len(values) - 1} linhas...")
update_count = 0

# Aplicar correções
for row_idx, row in enumerate(values[1:], 2):
    # Corrigir Forma de Pagamento
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'CONSTANTES'!{chr(65 + forma_col)}{row_idx}",
        valueInputOption="USER_ENTERED",
        body={"values": [["CARTÃO DE CRÉDITO"]]},
    ).execute()

    # Corrigir Tipo
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'CONSTANTES'!{chr(65 + tipo_col)}{row_idx}",
        valueInputOption="USER_ENTERED",
        body={"values": [["PAGAMENTO"]]},
    ).execute()

    update_count += 1
    if update_count % 10 == 0:
        print(f"  ✅ {update_count} linhas corrigidas...")

print(f"\n✅ Todas as {update_count} linhas foram corrigidas!")
print("\nNovas configurações:")
print("  • FORMA DE PAGAMENTO = 'CARTÃO DE CRÉDITO'")
print("  • TIPO = 'PAGAMENTO'")
print("\n✨ Pronto para executar main.py!")
