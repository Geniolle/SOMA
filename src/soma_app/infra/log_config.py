# src/soma_app/infra/log_config.py
from __future__ import annotations

import logging
import os
from datetime import datetime
from typing import Any, Dict, Optional, Tuple


def _pick_attr(obj: Any, names: list[str]) -> Optional[Any]:
    for n in names:
        if hasattr(obj, n):
            v = getattr(obj, n)
            if v is not None and v != "":
                return v
    return None


def _normalize_path(value: Any, default: str) -> str:
    if value is None:
        return default
    try:
        s = os.fspath(value)
        if isinstance(s, str):
            s = s.strip()
            return s or default
    except Exception:
        pass
    return default


def _resolve_dirs_from_settings(settings: Any) -> Tuple[str, str]:
    log_dir = _pick_attr(settings, ["LOG_DIR", "log_dir", "logs_dir", "LOGS_DIR"])
    artifacts_dir = _pick_attr(settings, ["ARTIFACTS_DIR", "artifacts_dir", "artifact_dir", "ARTIFACT_DIR"])

    log_dir = os.getenv("LOG_DIR") or log_dir
    artifacts_dir = os.getenv("ARTIFACTS_DIR") or artifacts_dir

    return _normalize_path(log_dir, "logs"), _normalize_path(artifacts_dir, "artifacts")


def ensure_artifacts_dirs(
    settings_or_log_dir: Any = None,
    *,
    log_dir: str = "logs",
    artifacts_dir: Optional[str] = None,
) -> Dict[str, str]:
    # settings object
    if settings_or_log_dir is not None and not isinstance(settings_or_log_dir, (str, bytes, os.PathLike)):
        resolved_log_dir, resolved_artifacts_dir = _resolve_dirs_from_settings(settings_or_log_dir)
        log_dir = resolved_log_dir
        artifacts_dir = resolved_artifacts_dir

    # direct path
    elif settings_or_log_dir is not None:
        log_dir = _normalize_path(settings_or_log_dir, "logs")

    log_dir = _normalize_path(os.getenv("LOG_DIR") or log_dir, "logs")
    artifacts_dir = _normalize_path(os.getenv("ARTIFACTS_DIR") or artifacts_dir, "artifacts")

    paths: Dict[str, str] = {
        "log_dir": log_dir,
        "artifacts_dir": artifacts_dir,
        "screenshots_dir": os.path.join(artifacts_dir, "screenshots"),
        "html_dir": os.path.join(artifacts_dir, "html"),
        "downloads_dir": os.path.join(artifacts_dir, "downloads"),
        "tmp_dir": os.path.join(artifacts_dir, "tmp"),
    }

    for p in paths.values():
        os.makedirs(p, exist_ok=True)

    return paths


def configure_logging(
    settings_or_log_dir: Any = "logs",
    *,
    app_level: Optional[str] = None,
    console_level: Optional[str] = None,
    audit_level: Optional[str] = None,
) -> Tuple[str, str]:
    """
    3 canais:
      - APP (root): técnico -> ficheiro soma_dev_*.log (+ console em WARNING por default)
      - AUDIT (soma_audit): auditoria estruturada -> ficheiro soma_audit_*.log
      - REPORT (soma_report): estilo legacy (sem timestamps) -> console + ficheiro soma_report_*.log

    Retorna (app_log_path, audit_log_path).
    """
    # Resolve dirs
    if settings_or_log_dir is not None and not isinstance(settings_or_log_dir, (str, bytes, os.PathLike)):
        log_dir, artifacts_dir = _resolve_dirs_from_settings(settings_or_log_dir)
        ensure_artifacts_dirs(log_dir=log_dir, artifacts_dir=artifacts_dir)
    else:
        log_dir = _normalize_path(settings_or_log_dir, "logs")
        ensure_artifacts_dirs(log_dir=log_dir)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    app_log_path = os.path.join(log_dir, f"soma_dev_{ts}.log")
    audit_log_path = os.path.join(log_dir, f"soma_audit_{ts}.log")
    report_log_path = os.path.join(log_dir, f"soma_report_{ts}.log")

    # Levels (ENV/params)
    app_level_name = (app_level or os.getenv("APP_LOG_LEVEL") or os.getenv("LOG_LEVEL") or "INFO").upper()
    # Para não poluir o console, default WARNING (como o script antigo que era “print”)
    console_level_name = (console_level or os.getenv("CONSOLE_LOG_LEVEL") or "WARNING").upper()
    audit_level_name = (audit_level or os.getenv("AUDIT_LOG_LEVEL") or "INFO").upper()
    report_level_name = (os.getenv("REPORT_LOG_LEVEL") or "INFO").upper()

    app_level_value = getattr(logging, app_level_name, logging.INFO)
    console_level_value = getattr(logging, console_level_name, logging.WARNING)
    audit_level_value = getattr(logging, audit_level_name, logging.INFO)
    report_level_value = getattr(logging, report_level_name, logging.INFO)

    # ---------------------------
    # APP / ROOT LOGGER
    # ---------------------------
    root = logging.getLogger()
    root.setLevel(app_level_value)

    app_fmt = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    for h in list(root.handlers):
        root.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    console = logging.StreamHandler()
    console.setLevel(console_level_value)
    console.setFormatter(app_fmt)

    app_file = logging.FileHandler(app_log_path, encoding="utf-8")
    app_file.setLevel(app_level_value)
    app_file.setFormatter(app_fmt)

    root.addHandler(console)
    root.addHandler(app_file)

    # ---------------------------
    # AUDIT LOGGER (ficheiro)
    # ---------------------------
    audit = logging.getLogger("soma_audit")
    audit.setLevel(audit_level_value)
    audit.propagate = False

    audit_fmt = logging.Formatter(
        fmt="%(asctime)s | AUDIT | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    for h in list(audit.handlers):
        audit.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    audit_file = logging.FileHandler(audit_log_path, encoding="utf-8")
    audit_file.setLevel(audit_level_value)
    audit_file.setFormatter(audit_fmt)
    audit.addHandler(audit_file)

    # ---------------------------
    # REPORT LOGGER (console + ficheiro, estilo legacy)
    # ---------------------------
    report = logging.getLogger("soma_report")
    report.setLevel(report_level_value)
    report.propagate = False

    report_fmt = logging.Formatter(fmt="%(message)s")

    for h in list(report.handlers):
        report.removeHandler(h)
        try:
            h.close()
        except Exception:
            pass

    report_console = logging.StreamHandler()
    report_console.setLevel(report_level_value)
    report_console.setFormatter(report_fmt)

    report_file = logging.FileHandler(report_log_path, encoding="utf-8")
    report_file.setLevel(report_level_value)
    report_file.setFormatter(report_fmt)

    report.addHandler(report_console)
    report.addHandler(report_file)

    # Arranque
    logging.getLogger(__name__).info("Logging ativo. Ficheiro: %s", app_log_path)
    logging.getLogger(__name__).info("Audit ativo. Ficheiro: %s", audit_log_path)
    logging.getLogger(__name__).info("Report legacy ativo. Ficheiro: %s", report_log_path)

    return app_log_path, audit_log_path