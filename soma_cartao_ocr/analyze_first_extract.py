#!/usr/bin/env python3
"""Analisa a primeira linha da sheet EXTRATO_CARTÃO sem escrever resultados."""

import json
import os
import re
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np
import yaml
from PIL import Image
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

try:
    from google.cloud import vision
except ImportError:
    vision = None


@dataclass
class Word:
    text: str
    x0: int
    y0: int
    x1: int
    y1: int
    confidence: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Movement:
    line: int
    data_movimento: str
    data_valor: str
    descricao: str
    pais: str
    moeda_original: str
    taxa_cambio: str
    debito_eur: str
    credito_eur: str
    confidence: float
    status: str
    motivos_revisao: str
    texto_ocr: str


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


def create_drive_service(credentials):
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def download_drive_image(drive_service, file_id: str, destination: Path) -> None:
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    with destination.open("wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()


def analyze_sheet_and_image():
    """Analisa a primeira linha sem escrever na sheet."""

    cfg_path = Path(__file__).with_name("config.yaml")
    cfg = load_config(cfg_path)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/cloud-platform",
    ]

    print("\n" + "="*100)
    print("🔍 ANÁLISE DA PRIMEIRA LINHA - MODO DRY RUN (Sem escrever na sheet)")
    print("="*100 + "\n")

    credentials = load_google_credentials(cfg, scopes)
    sheets_service = create_sheets_service(credentials)
    drive_service = create_drive_service(credentials)
    spreadsheet_id = cfg.get("google_sheets", {}).get("spreadsheet_id", "")

    # Obter primeira linha disponível
    response = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'EXTRATO_CARTÃO'!A1:Z500",
    ).execute()

    rows = response.get("values", [])
    if len(rows) < 2:
        print("❌ Nenhuma linha de dados encontrada!")
        return

    header = [str(h).strip().lower() for h in rows[0]]
    header_original = [str(h).strip() for h in rows[0]]

    img_col = next((idx for idx, c in enumerate(header) if "imagem" in c), 1)
    status_col = next((idx for idx, c in enumerate(header) if "status" in c), 2)
    nr_col = next((idx for idx, c in enumerate(header) if "extrato" in c), 0)

    # Encontrar primeira linha não processada
    first_available = None
    for idx in range(1, len(rows)):
        r = rows[idx]
        status_val = str(r[status_col]).strip() if len(r) > status_col else ""
        img_val = str(r[img_col]).strip() if len(r) > img_col else ""
        nr_val = str(r[nr_col]).strip() if len(r) > nr_col else ""
        if not status_val and img_val:
            first_available = (idx + 1, nr_val, img_val)
            break

    if not first_available:
        print("❌ Nenhuma linha disponível para processamento!")
        return

    linha_num, nr_extrato, img_path = first_available
    img_filename = Path(img_path).name

    print(f"📍 LINHA SELECIONADA:")
    print(f"   • Linha da sheet: {linha_num}")
    print(f"   • Número do extrato: {nr_extrato}")
    print(f"   • Caminho da imagem: {img_path}\n")

    # Baixar imagem
    print(f"📥 Baixando imagem do Google Drive...")

    drive_cfg = cfg.get("drive") or cfg.get("input", {}).get("drive", {})
    drive_cfg_temp = dict(drive_cfg)
    drive_cfg_temp["filename"] = img_filename

    # Localiazar imagem
    folder_id = drive_cfg_temp.get("folder_id", "")
    query = f"'{folder_id}' in parents and name = '{img_filename}' and trashed = false"

    response = drive_service.files().list(
        q=query,
        fields="files(id, name, mimeType, size, modifiedTime)",
        pageSize=1,
        spaces="drive",
        corpora="allDrives",
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        orderBy="modifiedTime desc",
    ).execute()

    files = response.get("files", [])
    if not files:
        print(f"❌ Imagem não encontrada: {img_filename}")
        return

    drive_info = files[0]

    with tempfile.TemporaryDirectory(prefix="soma_ocr_") as temp_dir:
        source_image = Path(temp_dir) / img_filename
        download_drive_image(drive_service, drive_info["id"], source_image)
        print(f"✅ Imagem baixada: {drive_info['name']} ({drive_info['size']} bytes)\n")

        # Mostrar informações da imagem
        img = cv2.imread(str(source_image))
        print(f"📏 Dimensões da imagem: {img.shape[1]}x{img.shape[0]} pixels\n")

        # Criar tabela resumida da análise (sem OCR pesado, apenas análise visual)
        print(f"{'='*100}")
        print(f"📊 ANÁLISE VISUAL DA IMAGEM")
        print(f"{'='*100}\n")

        print(f"{'Propriedade':<30} {'Valor':<70}")
        print(f"{'-'*100}")
        print(f"{'ID da imagem (Drive)':<30} {drive_info['id']:<70}")
        print(f"{'Nome do arquivo':<30} {drive_info['name']:<70}")
        print(f"{'Tamanho (bytes)':<30} {drive_info['size']:<70}")
        print(f"{'Tipo MIME':<30} {drive_info.get('mimeType', 'N/A'):<70}")
        print(f"{'Data de modificação':<30} {drive_info.get('modifiedTime', 'N/A'):<70}")
        print(f"{'Altura (pixels)':<30} {img.shape[0]:<70}")
        print(f"{'Largura (pixels)':<30} {img.shape[1]:<70}")
        print(f"{'Canais de cor':<30} {img.shape[2] if len(img.shape) > 2 else 1:<70}")
        print(f"{'-'*100}\n")

        print(f"✨ Análise concluída! Nenhum dado foi escrito na sheet.\n")
        print(f"💡 Próximos passos:")
        print(f"   1. Execute 'python main.py' para processar com OCR completo")
        print(f"   2. Os resultados serão salvos em 'output/resultado.json'")
        print(f"   3. A sheet CARTÃO será preenchida automaticamente\n")


if __name__ == "__main__":
    try:
        analyze_sheet_and_image()
    except Exception as e:
        print(f"ERRO: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
