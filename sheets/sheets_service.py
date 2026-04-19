from __future__ import annotations

import importlib
import os
import sys
from pathlib import Path
from typing import Any

import gspread


def _ensure_src_path() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    src_dir = base_dir / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))


def _load_settings() -> Any:
    _ensure_src_path()
    settings_cls = importlib.import_module("soma_app.config.settings").Settings
    env_file = (os.getenv("ENV_FILE") or "").strip() or None
    return settings_cls.from_env(env_file=env_file)


def obter_sheet():
    settings = _load_settings()
    client = gspread.service_account(filename=str(settings.google_credentials_path))
    spreadsheet = client.open_by_url(settings.spreadsheet_url)
    sheet = spreadsheet.worksheet(settings.sheet_contaordem)
    print(f"Worksheet '{settings.sheet_contaordem}' carregada com sucesso.")
    return sheet


def obter_todos_os_registros(sheet):
    registros = sheet.get_all_records()
    print(f"{len(registros)} linhas carregadas da folha.")
    return registros


def atualizar_linha(sheet, indice_linha, dados):
    cabecalho = sheet.row_values(1)
    mapa = {str(col).strip().upper(): i + 1 for i, col in enumerate(cabecalho)}

    for coluna, valor in dados.items():
        col_idx = mapa.get(str(coluna).strip().upper())
        if not col_idx:
            continue
        sheet.update_cell(indice_linha, col_idx, valor)

    print(f"Linha {indice_linha} atualizada com sucesso.")
