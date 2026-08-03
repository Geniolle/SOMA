#!/usr/bin/env python3
"""Verifica e corrige datas invertidas nas sheets."""

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

print("📅 Verificando datas nas sheets CARTÃO e SAÍDAS...\n")

def check_and_fix_dates(sheet_name, date_columns):
    """Verifica e corrige datas invertidas em uma sheet."""
    print(f"Sheet '{sheet_name}':")

    try:
        response = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1:Z5000",
        ).execute()

        values = response.get("values", [])
        if not values:
            print(f"  ❌ Sheet vazia\n")
            return

        header = values[0]
        header_lower = [h.lower() for h in header]

        # Encontrar colunas de data
        col_indices = {}
        for col_name in date_columns:
            col_idx = next((i for i, h in enumerate(header_lower) if col_name.lower() in h), None)
            if col_idx is not None:
                col_indices[col_name] = col_idx
                print(f"  Coluna '{col_name}': {header[col_idx]}")

        if not col_indices:
            print(f"  ⚠️  Nenhuma coluna de data encontrada\n")
            return

        print()

        # Verificar datas
        inverted_dates = []
        date_pattern = re.compile(r"^(\d{2})[/\-](\d{2})$")

        for row_idx, row in enumerate(values[1:], 2):
            for col_name, col_idx in col_indices.items():
                if col_idx >= len(row):
                    continue

                date_str = str(row[col_idx]).strip()
                if not date_str:
                    continue

                match = date_pattern.match(date_str)
                if not match:
                    continue

                day = int(match.group(1))
                month = int(match.group(2))

                # Verificar se é data invertida
                # Data invertida: dia > 12 e mês <= 12 (significa que pode estar trocado)
                # Ou: mês > 12 (mês inválido)
                if month > 12:
                    print(f"  ⚠️  Linha {row_idx}: '{col_name}' = '{date_str}' - MÊS INVÁLIDO (>{12})")
                    # Supor que foi invertida
                    corrected = f"{month:02d}/{day:02d}"
                    inverted_dates.append((row_idx, col_idx, col_name, sheet_name, date_str, corrected))
                    print(f"      → Sugestão de correção: {corrected}")

                elif day > 12 and month <= 12:
                    # Potencial inversão
                    print(f"  ⓘ Linha {row_idx}: '{col_name}' = '{date_str}' - POSSÍVEL INVERSÃO (dia>{day}, mês<={month})")
                    # Verificar contexto
                    if day <= 31 and month <= 12:  # Se a inversão fizer sentido
                        corrected = f"{month:02d}/{day:02d}"
                        print(f"      → Se invertida, ficaria: {corrected}")

        if not inverted_dates:
            print("  ✅ Todas as datas estão válidas (formato DD/MM com dia <= 31, mês <= 12)")
        else:
            print(f"\n  ⚠️  {len(inverted_dates)} data(s) com mês inválido encontrada(s)")

            # Perguntar se quer corrigir
            if input("\n  Deseja corrigir as datas? (s/n): ").lower() == "s":
                for row_idx, col_idx, col_name, sheet, old_date, new_date in inverted_dates:
                    col_letter = chr(65 + col_idx)
                    sheets_service.spreadsheets().values().update(
                        spreadsheetId=spreadsheet_id,
                        range=f"'{sheet}'!{col_letter}{row_idx}",
                        valueInputOption="USER_ENTERED",
                        body={"values": [[new_date]]},
                    ).execute()
                    print(f"    ✅ {sheet}!{col_letter}{row_idx}: {old_date} → {new_date}")

        print()

    except Exception as e:
        print(f"  ❌ Erro ao acessar sheet: {e}\n")


# Verificar ambas as sheets
check_and_fix_dates("CARTÃO", ["data mov", "data valor"])
check_and_fix_dates("SAÍDAS", ["data"])

print("✅ Verificação concluída!")
