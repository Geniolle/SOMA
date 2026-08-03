#!/usr/bin/env python3
"""Verifica linhas disponíveis para processamento na sheet EXTRATO_CARTÃO."""

import json
import sys
from pathlib import Path

import yaml
from google.oauth2 import service_account
from googleapiclient.discovery import build


def load_config(path: Path) -> dict:
    if not path.exists():
        raise FileNotFoundError(f"Ficheiro de configuração não encontrado: {path}")
    with path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle) or {}


def load_google_credentials(cfg: dict, scopes: list) -> service_account.Credentials:
    creds_path_str = cfg.get("google", {}).get("service_account_file", "").strip()
    if not creds_path_str:
        raise ValueError("Caminho do ficheiro de conta de serviço não configurado.")

    creds_path = Path(creds_path_str)
    if not creds_path.is_absolute():
        creds_path = (Path(__file__).parent / creds_path).resolve()

    if not creds_path.exists():
        raise FileNotFoundError(f"Ficheiro de credenciais não encontrado: {creds_path}")

    return service_account.Credentials.from_service_account_file(
        str(creds_path), scopes=scopes
    )


def create_sheets_service(credentials):
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def check_extrato_status(sheets_service, spreadsheet_id, worksheet: str = "EXTRATO_CARTÃO"):
    """Verifica e exibe as linhas disponíveis para processamento."""

    print(f"\n{'='*80}")
    print(f"📋 VERIFICAÇÃO DA SHEET: {worksheet}")
    print(f"{'='*80}\n")

    try:
        response = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{worksheet}'!A1:Z500",
        ).execute()

        rows = response.get("values", [])
        if not rows:
            print("❌ Sheet vazia ou não encontrada!")
            return

        # Extrair cabeçalho
        header = [str(h).strip().lower() for h in rows[0]]
        print(f"📌 CABEÇALHO ENCONTRADO ({len(header)} colunas):")
        print(f"   {rows[0]}\n")

        # Encontrar colunas importantes
        img_col = next((idx for idx, c in enumerate(header) if "imagem" in c), 1)
        status_col = next((idx for idx, c in enumerate(header) if "status" in c), 2)
        nr_col = next((idx for idx, c in enumerate(header) if "extrato" in c), 0)

        print(f"🔍 COLUNAS DETECTADAS:")
        print(f"   • Coluna IMAGEM: índice {img_col} ({rows[0][img_col] if img_col < len(rows[0]) else 'N/A'})")
        print(f"   • Coluna STATUS: índice {status_col} ({rows[0][status_col] if status_col < len(rows[0]) else 'N/A'})")
        print(f"   • Coluna EXTRATO: índice {nr_col} ({rows[0][nr_col] if nr_col < len(rows[0]) else 'N/A'})\n")

        # Processar linhas
        unprocessed = []
        processed = []
        empty = []

        for idx in range(1, len(rows)):
            r = rows[idx]
            status_val = str(r[status_col]).strip() if len(r) > status_col else ""
            img_val = str(r[img_col]).strip() if len(r) > img_col else ""
            nr_val = str(r[nr_col]).strip() if len(r) > nr_col else ""

            row_num = idx + 1

            if not img_val:
                # Linha vazia (sem imagem)
                empty.append({
                    "linha": row_num,
                    "numero_extrato": nr_val,
                    "imagem": img_val,
                    "status": status_val
                })
            elif not status_val:
                # Linha disponível para processamento
                unprocessed.append({
                    "linha": row_num,
                    "numero_extrato": nr_val,
                    "imagem": img_val,
                    "status": status_val
                })
            else:
                # Linha já processada
                processed.append({
                    "linha": row_num,
                    "numero_extrato": nr_val,
                    "imagem": img_val,
                    "status": status_val
                })

        # Exibir resultados
        print(f"{'='*80}")
        print(f"📊 RESUMO DE LINHAS:\n")

        print(f"✅ DISPONÍVEIS PARA PROCESSAMENTO (Status = vazio): {len(unprocessed)}")
        if unprocessed:
            print(f"   {'-'*76}")
            for item in unprocessed:
                print(f"   Linha {item['linha']:3d} | Extrato: {item['numero_extrato']:15s} | Imagem: {item['imagem']}")
            print(f"   {'-'*76}")
        else:
            print("   ℹ️  Nenhuma linha disponível")

        print(f"\n⏳ JÁ PROCESSADAS (Status ≠ vazio): {len(processed)}")
        if processed:
            print(f"   {'-'*76}")
            for item in processed:
                print(f"   Linha {item['linha']:3d} | Extrato: {item['numero_extrato']:15s} | Status: {item['status']}")
            print(f"   {'-'*76}")
        else:
            print("   ℹ️  Nenhuma linha processada")

        print(f"\n⚪ LINHAS VAZIAS (sem imagem): {len(empty)}")
        if len(empty) <= 5:
            for item in empty:
                if item['numero_extrato']:
                    print(f"   Linha {item['linha']:3d} | Extrato: {item['numero_extrato']:15s}")

        print(f"\n{'='*80}")
        print(f"📈 TOTAL DE LINHAS: {len(rows) - 1} (excluindo cabeçalho)")
        print(f"{'='*80}\n")

        # Mostrar próximo extrato a ser processado
        if unprocessed:
            next_item = unprocessed[0]
            print(f"⏭️  PRÓXIMO EXTRATO A PROCESSAR:")
            print(f"   Linha: {next_item['linha']}")
            print(f"   Número: {next_item['numero_extrato']}")
            print(f"   Imagem: {next_item['imagem']}\n")

    except Exception as e:
        print(f"❌ Erro ao consultar sheet: {e}")
        raise


def main():
    cfg_path = Path(__file__).with_name("config.yaml")
    cfg = load_config(cfg_path)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
    ]

    credentials = load_google_credentials(cfg, scopes)
    sheets_service = create_sheets_service(credentials)
    spreadsheet_id = cfg.get("google_sheets", {}).get("spreadsheet_id", "")

    if not spreadsheet_id:
        raise ValueError("Defina google_sheets.spreadsheet_id no config.yaml")

    check_extrato_status(sheets_service, spreadsheet_id)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERRO: {e}", file=sys.stderr)
        sys.exit(1)
