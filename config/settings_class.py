from __future__ import annotations

import importlib
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
    module = importlib.import_module("soma_app.config.settings")
    return module.Settings


Settings = _load_settings()

__all__ = ["Settings"]
