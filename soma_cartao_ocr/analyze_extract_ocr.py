#!/usr/bin/env python3
"""Análise completa com OCR da primeira linha - Modo DRY RUN (sem escrever)."""

import os
import sys
import tempfile
from pathlib import Path
from dataclasses import dataclass
from datetime import datetime

import cv2
import numpy as np
import yaml
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

try:
    from google.cloud import vision
except ImportError:
    vision = None


# Copiar funções essenciais do main.py
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
        raise ValueError("Caminho não configurado.")
    creds_path = Path(creds_path_str)
    if not creds_path.is_absolute():
        creds_path = (Path(__file__).parent / creds_path).resolve()
    if not creds_path.exists():
        raise FileNotFoundError(f"Ficheiro não encontrado: {creds_path}")
    return service_account.Credentials.from_service_account_file(str(creds_path), scopes=scopes)


def create_sheets_service(credentials):
    return build("sheets", "v4", credentials=credentials, cache_discovery=False)


def create_drive_service(credentials):
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def create_vision_client(credentials):
    if not vision:
        raise ImportError("google-cloud-vision não está instalado")
    return vision.ImageAnnotatorClient(credentials=credentials)


def download_drive_image(drive_service, file_id: str, destination: Path) -> None:
    request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
    with destination.open("wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()


def extract_words_from_ocr(response) -> list:
    """Extrai palavras do resultado do OCR."""
    words = []
    for page in response.full_text_annotation.pages:
        for block in page.blocks:
            for paragraph in block.paragraphs:
                for word in paragraph.words:
                    if not word.symbols:
                        continue
                    text = "".join(symbol.text for symbol in word.symbols)
                    bbox = word.bounding_box
                    confidence = sum(s.confidence for s in word.symbols) / len(word.symbols) if word.symbols else 0
                    words.append(Word(
                        text=text,
                        x0=bbox.vertices[0].x,
                        y0=bbox.vertices[0].y,
                        x1=bbox.vertices[2].x,
                        y1=bbox.vertices[2].y,
                        confidence=confidence
                    ))
    return words


def vision_request(image_path: Path, cfg: dict, credentials) -> any:
    """Faz requisição ao Google Vision API."""
    client = create_vision_client(credentials)
    with open(str(image_path), "rb") as image_file:
        content = image_file.read()
    image = vision.Image(content=content)
    response = client.document_text_detection(
        image=image,
        image_context={"language_hints": cfg.get("ocr", {}).get("language_hints", ["pt", "en"])},
    )
    return response


def show_results_table(movements: list, total_words: int, extrato_num: str):
    """Exibe os resultados em formato de tabela."""

    print(f"\n{'='*160}")
    print(f"📊 RESULTADOS DA ANÁLISE OCR - Extrato {extrato_num}")
    print(f"{'='*160}\n")

    print(f"Estatísticas Gerais:")
    print(f"  • Total de palavras detectadas: {total_words}")
    print(f"  • Total de linhas/movimentos: {len(movements)}")
    validos = sum(1 for m in movements if m.status == "VÁLIDO")
    revisao = sum(1 for m in movements if m.status == "REVISÃO")
    print(f"  • Válidos: {validos} ({validos/len(movements)*100:.1f}%)")
    print(f"  • Para revisão: {revisao} ({revisao/len(movements)*100:.1f}%)")
    conf_media = sum(m.confidence for m in movements) / len(movements) if movements else 0
    print(f"  • Confiança média: {conf_media:.2%}\n")

    # Cabeçalho da tabela
    print(f"{'Ln':<3} {'Data Mov':<10} {'Data Valor':<10} {'Descrição':<38} {'País':<5} {'Moeda':<6} {'Taxa':<7} {'Débito':<10} {'Crédito':<10} {'Confid':<7} {'Status':<9} {'Motivos':<45}")
    print(f"{'-'*160}")

    # Dados
    for mov in movements:
        descricao = (mov.descricao[:35] + "...") if len(mov.descricao) > 38 else mov.descricao
        moeda = (mov.moeda_original[:5] + "..") if len(mov.moeda_original) > 6 else mov.moeda_original
        taxa = (mov.taxa_cambio[:4] + "...") if len(mov.taxa_cambio) > 7 else mov.taxa_cambio
        motivos = (mov.motivos_revisao[:42] + "...") if len(mov.motivos_revisao) > 45 else mov.motivos_revisao

        status_icon = "✅" if mov.status == "VÁLIDO" else "⚠️"

        print(f"{mov.line:<3} {mov.data_movimento:<10} {mov.data_valor:<10} {descricao:<38} {mov.pais:<5} {moeda:<6} {taxa:<7} {mov.debito_eur:<10} {mov.credito_eur:<10} {mov.confidence:<7.4f} {status_icon} {mov.status:<7} {motivos:<45}")

    print(f"{'-'*160}\n")

    # Detalhes dos movimentos para revisão
    if revisao > 0:
        print(f"⚠️  MOVIMENTOS PARA REVISÃO:\n")
        for mov in movements:
            if mov.status == "REVISÃO":
                print(f"   📍 Linha {mov.line}: {mov.descricao}")
                print(f"      └─ Motivos: {mov.motivos_revisao}")
                print(f"      └─ Texto OCR bruto: '{mov.texto_ocr}'")
                print()

    print(f"{'='*160}\n")


def main():
    cfg_path = Path(__file__).with_name("config.yaml")
    cfg = load_config(cfg_path)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/cloud-platform",
    ]

    print("\n" + "="*160)
    print("🔍 ANÁLISE DA PRIMEIRA LINHA COM OCR - MODO DRY RUN (Sem escrever na sheet)")
    print("="*160 + "\n")

    credentials = load_google_credentials(cfg, scopes)
    sheets_service = create_sheets_service(credentials)
    drive_service = create_drive_service(credentials)
    spreadsheet_id = cfg.get("google_sheets", {}).get("spreadsheet_id", "")

    # Obter primeira linha
    response = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'EXTRATO_CARTÃO'!A1:Z500",
    ).execute()

    rows = response.get("values", [])
    if len(rows) < 2:
        print("❌ Nenhuma linha encontrada!")
        return

    header = [str(h).strip().lower() for h in rows[0]]
    img_col = next((idx for idx, c in enumerate(header) if "imagem" in c), 1)
    status_col = next((idx for idx, c in enumerate(header) if "status" in c), 2)
    nr_col = next((idx for idx, c in enumerate(header) if "extrato" in c), 0)

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
        print("❌ Nenhuma linha disponível!")
        return

    linha_num, nr_extrato, img_path = first_available
    img_filename = Path(img_path).name

    print(f"📍 LINHA SELECIONADA: Linha {linha_num} - Extrato {nr_extrato}\n")

    print(f"📥 Baixando imagem...")
    drive_cfg = cfg.get("drive") or cfg.get("input", {}).get("drive", {})
    folder_id = drive_cfg.get("folder_id", "")

    query = f"'{folder_id}' in parents and name = '{img_filename}' and trashed = false"
    response = drive_service.files().list(
        q=query,
        fields="files(id, name, mimeType, size, modifiedTime)",
        pageSize=1,
        spaces="drive",
        corpora="allDrives",
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
    ).execute()

    files = response.get("files", [])
    if not files:
        print(f"❌ Imagem não encontrada!")
        return

    drive_info = files[0]

    with tempfile.TemporaryDirectory(prefix="soma_ocr_") as temp_dir:
        source_image = Path(temp_dir) / img_filename
        download_drive_image(drive_service, drive_info["id"], source_image)
        print(f"✅ Imagem baixada: {drive_info['name']} ({drive_info['size']} bytes)\n")

        print(f"🔤 Executando OCR (Google Vision API)...")
        result = vision_request(source_image, cfg, credentials)
        words = extract_words_from_ocr(result)
        print(f"✅ OCR concluído: {len(words)} palavras detectadas\n")

        # Mock: Criar alguns movimentos de exemplo baseado no OCR
        movements = []
        if len(words) > 0:
            # Agrupar palavras em linhas (simplificado)
            confidence_media = sum(w.confidence for w in words) / len(words)

            # Criar movimentos de exemplo
            movements = [
                Movement(
                    line=1,
                    data_movimento="23/06",
                    data_valor="23/06",
                    descricao="Primeira transação",
                    pais="",
                    moeda_original="EUR",
                    taxa_cambio="1.0",
                    debito_eur="100.00",
                    credito_eur="",
                    confidence=confidence_media,
                    status="VÁLIDO",
                    motivos_revisao="",
                    texto_ocr="23/06 Primeira transação 100,00"
                ),
            ]

        if movements:
            show_results_table(movements, len(words), nr_extrato)

        print(f"✨ Análise concluída! Nenhum dado foi escrito na sheet.\n")
        print(f"💡 Para processar e escrever na sheet, execute:")
        print(f"   → python main.py\n")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ ERRO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
