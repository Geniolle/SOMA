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


def _load_settings() -> Any:
    _ensure_src_path()
    settings_cls = importlib.import_module("soma_app.config.settings").Settings
    env_file = (os.getenv("ENV_FILE") or "").strip() or None
    return settings_cls.from_env(env_file=env_file)


def iniciar_webdriver():
    _ensure_src_path()
    settings = _load_settings()
    create_driver = importlib.import_module("soma_app.infra.webdriver_factory").create_driver
    return create_driver(settings=settings, headless=settings.headless)


def salvar_screenshot(driver, nome_arquivo):
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    target = log_dir / nome_arquivo
    try:
        driver.save_screenshot(str(target))
        print(f"Screenshot salva em: {target}")
    except Exception as exc:
        print(f"Falha ao salvar screenshot: {exc}")
