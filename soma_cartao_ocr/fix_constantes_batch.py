#!/usr/bin/env python3
"""Corrige sheet CONSTANTES usando batch update."""

from pathlib import Path
import time

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

print("📊 Corrigindo sheet CONSTANTES com batch update...\n")

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

print(f"Preparando correção de {len(values) - 1} linhas...\n")

# Preparar batch update
data_to_update = []

for row_idx in range(2, len(values) + 1):
    # Forma de Pagamento
    data_to_update.append({
        "range": f"'CONSTANTES'!{chr(65 + forma_col)}{row_idx}",
        "values": [["CARTÃO DE CRÉDITO"]]
    })

    # Tipo
    data_to_update.append({
        "range": f"'CONSTANTES'!{chr(65 + tipo_col)}{row_idx}",
        "values": [["PAGAMENTO"]]
    })

print(f"Enviando {len(data_to_update)} updates em batch...\n")

# Executar batch update
for i in range(0, len(data_to_update), 30):  # 30 updates por batch
    batch = data_to_update[i:i+30]

    body = {
        "data": batch,
        "valueInputOption": "USER_ENTERED"
    }

    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()

    print(f"✅ {min(i + 30, len(data_to_update))}/{len(data_to_update)} updates completados")

    # Pequeno delay para respeitar rate limit
    if i + 30 < len(data_to_update):
        time.sleep(0.5)

print(f"\n✅ Sheet CONSTANTES corrigida com sucesso!")
print("\nNovas configurações:")
print("  • FORMA DE PAGAMENTO = 'CARTÃO DE CRÉDITO'")
print("  • TIPO = 'PAGAMENTO'")
print("\n🚀 Pronto para executar main.py!")
