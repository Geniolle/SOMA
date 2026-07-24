from __future__ import annotations

import os


def env_bool(name: str, default: bool = False) -> bool:
    v = (os.getenv(name) or "").strip().lower()
    if v in {"1", "true", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "no", "n", "off"}:
        return False
    return default


def env_str(name: str, default: str = "") -> str:
    v = (os.getenv(name) or "").strip()
    return v if v else default


def env_int(name: str, default: int) -> int:
    v = (os.getenv(name) or "").strip()
    try:
        return int(v)
    except (TypeError, ValueError):
        return default
