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


def atualizar_caixas_soma(driver):
    _ensure_src_path()
    settings = _load_settings()

    actions_module = importlib.import_module("soma_app.automation.actions")
    sheets_module = importlib.import_module("soma_app.infra.sheets_client")
    workflow_module = importlib.import_module("soma_app.workflows.process_caixas_bancos")

    actions = actions_module.Actions(
        driver,
        actions_module.ActionConfig(
            timeout_seconds=settings.timeout_seconds,
            screenshots_dir=settings.screenshots_dir,
        ),
    )
    sheets = sheets_module.SheetsClient(settings)
    workflow_module.atualizar_caixas_bancos(sheets, actions, settings)
