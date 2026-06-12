from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

from selenium.webdriver.common.by import By

Locator = Tuple[str, str]

DEFAULT_LOCATORS_PATH = Path(__file__).with_name("locators.json")
PROJECT_ROOT = Path(__file__).resolve().parents[3]

_BY_ALIASES = {
    "xpath": By.XPATH,
    "css": By.CSS_SELECTOR,
    "css selector": By.CSS_SELECTOR,
    "id": By.ID,
    "name": By.NAME,
    "class": By.CLASS_NAME,
    "class_name": By.CLASS_NAME,
    "tag": By.TAG_NAME,
    "tag_name": By.TAG_NAME,
    "link_text": By.LINK_TEXT,
    "partial_link_text": By.PARTIAL_LINK_TEXT,
}


def _is_locator(value: Any) -> bool:
    return isinstance(value, tuple) and len(value) == 2 and all(isinstance(item, str) for item in value)


def _is_locator_list(value: Any) -> bool:
    return isinstance(value, list) and all(_is_locator(item) for item in value)


def _resolve_path(raw_path: Any) -> Path:
    raw = str(raw_path or "").strip()
    if not raw:
        return DEFAULT_LOCATORS_PATH

    path = Path(raw)
    if path.is_absolute():
        return path

    cwd_candidate = (Path.cwd() / path).resolve()
    if cwd_candidate.exists():
        return cwd_candidate

    return (PROJECT_ROOT / path).resolve()


def _coerce_locator(raw: Any, default: Locator | None) -> Locator | None:
    if isinstance(raw, str) and raw.strip():
        return (By.XPATH, raw.strip())

    if isinstance(raw, dict):
        by_raw = str(raw.get("by", "xpath")).strip().lower()
        value = str(raw.get("value", "")).strip()
        by = _BY_ALIASES.get(by_raw)
        if by and value:
            return (by, value)

    return default


def _coerce_locator_list(raw: Any, default: Iterable[Locator]) -> list[Locator]:
    if isinstance(raw, list):
        out: list[Locator] = []
        for item in raw:
            loc = _coerce_locator(item, None)
            if loc is not None:
                out.append(loc)
        if out:
            return out

    loc = _coerce_locator(raw, None)
    if loc is not None:
        return [loc]

    return list(default)


def load_page_locator_config(settings: Any, page_name: str) -> Dict[str, Any]:
    raw_path = (
        getattr(settings, "locators_path", None)
        or getattr(settings, "LOCATORS_PATH", None)
        or os.getenv("LOCATORS_PATH")
        or DEFAULT_LOCATORS_PATH
    )
    path = _resolve_path(raw_path)
    if not path.exists():
        return {}

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"Locator config invalido: {path} | {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError(f"Locator config invalido: {path} | raiz deve ser um objeto JSON.")

    page_cfg = data.get(page_name, {})
    if page_cfg is None:
        return {}
    if not isinstance(page_cfg, dict):
        raise ValueError(f"Locator config invalido: {path} | secao '{page_name}' deve ser um objeto.")
    return page_cfg


def apply_locator_overrides(page_obj: Any, page_name: str) -> None:
    cfg = load_page_locator_config(getattr(page_obj, "settings", None), page_name)
    if not cfg:
        return

    cls = type(page_obj)
    tuple_overrides: dict[Locator, Locator] = {}
    explicit_lists: set[str] = set()

    for name, default in vars(cls).items():
        if not name.isupper():
            continue

        current = getattr(page_obj, name, default)

        if _is_locator(default):
            loc = _coerce_locator(cfg.get(name), current) if name in cfg else current
            setattr(page_obj, name, loc)
            tuple_overrides[default] = loc
            continue

        if _is_locator_list(default):
            if name in cfg:
                setattr(page_obj, name, _coerce_locator_list(cfg.get(name), current))
                explicit_lists.add(name)

    for name, default in vars(cls).items():
        if not name.isupper() or name in explicit_lists or not _is_locator_list(default):
            continue

        current = list(getattr(page_obj, name, default))
        updated = [tuple_overrides.get(loc, loc) for loc in current]
        if updated != current:
            setattr(page_obj, name, updated)
