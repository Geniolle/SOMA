from __future__ import annotations

import importlib
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


def ensure_src_path() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    src_dir = base_dir / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))


def load_settings() -> Any:
    ensure_src_path()
    settings_cls = importlib.import_module("soma_app.config.settings").Settings
    env_file = (os.getenv("ENV_FILE") or "").strip() or None
    return settings_cls.from_env(env_file=env_file)


def build_actions(driver, settings) -> Any:
    actions_module = importlib.import_module("soma_app.automation.actions")
    return actions_module.Actions(
        driver,
        actions_module.ActionConfig(
            timeout_seconds=settings.timeout_seconds,
            screenshots_dir=settings.screenshots_dir,
        ),
    )


def to_row(index: int, linha: dict[str, Any], tipo_forcado: str | None = None):
    ensure_src_path()
    models = importlib.import_module("soma_app.domain.models")
    raw = dict(linha)
    if tipo_forcado:
        raw["TIPO"] = tipo_forcado
    return models.ContaOrdemRow.from_table_row(row_number=index, raw=raw)


def update_sheet_fields(sheet, index: int, values: dict[str, Any]) -> None:
    header = sheet.row_values(1)
    normalized = {str(col).strip().upper(): i + 1 for i, col in enumerate(header)}

    for field, value in values.items():
        col = normalized.get(field.strip().upper())
        if not col:
            continue
        sheet.update_cell(index, col, value)


def build_default_updates(*, doc_soma: str, dados_doc: str | None = None) -> dict[str, Any]:
    updates: dict[str, Any] = {
        "DOC. SOMA": doc_soma,
        "IDUSER": "USERJOB",
        "TIMESTAMP": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }
    if dados_doc:
        updates["DADOS DOC"] = dados_doc
    return updates
