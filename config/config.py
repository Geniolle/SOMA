from __future__ import annotations

import importlib
import os
import sys
from pathlib import Path
from typing import Any


def _ensure_src_path() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    src_dir = base_dir / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))


def _load_settings_class() -> Any:
    _ensure_src_path()
    return importlib.import_module("soma_app.config.settings").Settings


def _resolve_env_file() -> str | None:
    env_file = (os.getenv("ENV_FILE") or "").strip()
    if env_file:
        return env_file

    base_dir = Path(__file__).resolve().parent.parent
    default_env = base_dir / "deploy" / ".env"
    if default_env.exists():
        return str(default_env)
    return None


Settings = _load_settings_class()
settings = Settings.from_env(env_file=_resolve_env_file())

EMAIL_SOMA = settings.site_user
SENHA_SOMA = settings.site_password
SOMA_URL_INICIAL = settings.site_login_url
HEADLESS = settings.headless
SELENIUM_TIMEOUT = settings.timeout_seconds
CREDENCIAIS_PATH = str(settings.google_credentials_path)
SPREADSHEET_URL = settings.spreadsheet_url
SHEET_NAME = settings.sheet_contaordem
