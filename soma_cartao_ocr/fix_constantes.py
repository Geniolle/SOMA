#!/usr/bin/env python3
"""Verifica e corrige o mapeamento na sheet CONSTANTES."""

import json
import os
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build

# Carregar configuração
import yaml

config_path = Path(__file__).parent / "config.yaml"
with open(config_path, encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

# Carregar credenciais
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

print("📊 Verificando sheet CONSTANTES...")
print(f"Spreadsheet ID: {spreadsheet_id}\n")

# Ler dados da sheet CONSTANTES
try:
    response = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range="'CONSTANTES'!A1:Z5000",
    ).execute()

    values = response.get("values", [])
    if not values:
        print("❌ Sheet CONSTANTES está vazia!")
        exit(1)

    print(f"✅ Sheet CONSTANTES encontrada ({len(values)} linhas)\n")

    # Cabeçalho
    header = values[0]
    print(f"Cabeçalho: {header}\n")

    # Encontrar colunas
    header_lower = [h.lower() for h in header]

    texto_col = next((i for i, h in enumerate(header_lower) if "texto" in h), None)
    forma_col = next((i for i, h in enumerate(header_lower) if "forma de pagamento" in h), None)
    tipo_col = next((i for i, h in enumerate(header_lower) if h == "tipo"), None)

    print(f"Coluna 'Texto': {texto_col} ({header[texto_col] if texto_col is not None else 'NÃO ENCONTRADA'})")
    print(f"Coluna 'Forma de Pagamento': {forma_col} ({header[forma_col] if forma_col is not None else 'NÃO ENCONTRADA'})")
    print(f"Coluna 'Tipo': {tipo_col} ({header[tipo_col] if tipo_col is not None else 'NÃO ENCONTRADA'})\n")

    if forma_col is None or tipo_col is None:
        print("❌ Colunas 'Forma de Pagamento' ou 'Tipo' não encontradas!")
        print(f"Colunas disponíveis: {header}")
        exit(1)

    # Mostrar dados atuais
    print("📝 Dados atuais na sheet CONSTANTES:")
    print(f"{'Linha':<6} {'Texto':<30} {'Forma Pagamento':<20} {'Tipo':<15}")
    print("-" * 75)

    rows_to_update = []
    for i, row in enumerate(values[1:], 1):
        texto = str(row[texto_col]).strip() if len(row) > texto_col else ""
        forma = str(row[forma_col]).strip() if len(row) > forma_col else ""
        tipo = str(row[tipo_col]).strip() if len(row) > tipo_col else ""

        if texto:
            print(f"{i:<6} {texto:<30} {forma:<20} {tipo:<15}")

            # Verificar se precisa corrigir
            if forma != "CARTÃO DE CRÉDITO" or tipo != "PAGAMENTO":
                rows_to_update.append((i, texto, forma, tipo))

    print()

    if rows_to_update:
        print(f"⚠️  {len(rows_to_update)} linha(s) precisam ser corrigidas:\n")

        # Mostrar o que será corrigido
        for row_idx, texto, forma, tipo in rows_to_update:
            print(f"Linha {row_idx}: {texto}")
            print(f"  Forma Pagamento: '{forma}' → 'CARTÃO DE CRÉDITO'")
            print(f"  Tipo: '{tipo}' → 'PAGAMENTO'")
            print()

        # Perguntar para o usuário
        confirm = input("Deseja aplicar as correções? (s/n): ").lower()

        if confirm == "s":
            # Aplicar correções
            updates = []
            for row_idx, _, _, _ in rows_to_update:
                # Forma de Pagamento
                updates.append({
                    "range": f"'CONSTANTES'!{chr(65 + forma_col)}{row_idx + 1}",
                    "values": [["CARTÃO DE CRÉDITO"]]
                })
                # Tipo
                updates.append({
                    "range": f"'CONSTANTES'!{chr(65 + tipo_col)}{row_idx + 1}",
                    "values": [["PAGAMENTO"]]
                })

            # Executar atualizações em batch
            for update in updates:
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=update["range"],
                    valueInputOption="USER_ENTERED",
                    body={"values": update["values"]},
                ).execute()
                print(f"✅ {update['range']} atualizado")

            print("\n✅ Correções aplicadas com sucesso!")
        else:
            print("Operação cancelada.")
    else:
        print("✅ Sheet CONSTANTES já está correta!")
        print("Todos os registros têm:")
        print("  • Forma de Pagamento: CARTÃO DE CRÉDITO")
        print("  • Tipo: PAGAMENTO")

except Exception as e:
    print(f"❌ Erro ao acessar sheet: {e}")
    import traceback
    traceback.print_exc()
    exit(1)
