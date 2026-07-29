from __future__ import annotations

import os
import tempfile
from pathlib import Path


def _ensure_test_temp_dir() -> str:
    root = Path(__file__).resolve().parents[1]
    temp_dir = root / ".pytest-tmp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    return str(temp_dir)


_TEST_TEMP_DIR = _ensure_test_temp_dir()


def pytest_configure(config) -> None:  # type: ignore[override]
    os.environ.setdefault("TMPDIR", _TEST_TEMP_DIR)
    os.environ.setdefault("TEMP", _TEST_TEMP_DIR)
    os.environ.setdefault("TMP", _TEST_TEMP_DIR)
    tempfile.tempdir = _TEST_TEMP_DIR
