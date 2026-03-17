# src/soma_app/infra/trace.py
from __future__ import annotations

import logging
import time
import uuid
from contextlib import contextmanager
from typing import Any, Dict, Iterator, Optional, Union


def new_run_id(n: int = 10) -> str:
    """
    Gera um run_id curto (hex), compatível com o código existente.
    Default: 10 chars (ex: a873734fa0).
    """
    if n <= 0:
        n = 10
    return uuid.uuid4().hex[:n]


def fmt_kv(fields: Dict[str, Any]) -> str:
    """
    Formata um dict em estilo 'k=v | k=v', ignorando None.
    """
    parts = []
    for k, v in fields.items():
        if v is None:
            continue
        parts.append(f"{k}={v}")
    return " | ".join(parts)


def _resolve_level(level: Union[int, str]) -> int:
    if isinstance(level, str):
        return getattr(logging, level.upper(), logging.INFO)
    return int(level)


def log_kv(
    logger: Optional[logging.Logger],
    msg: str,
    *args: Any,
    level: Union[int, str] = "INFO",
    **fields: Any,
) -> None:
    """
    Log com key-values no padrão usado no projeto:
      <msg> | k=v | k=v

    Compatível com chamadas antigas que passavam um 3º argumento posicional:
      - log_kv(logger, "msg", "INFO", a=1)
      - log_kv(logger, "msg", {"a": 1, "b": 2})
      - log_kv(logger, "msg", "ERROR", {"a": 1})
    """
    if logger is None:
        logger = logging.getLogger(__name__)

    lvl: Union[int, str] = level

    # Aceita até 2 args posicionais "extras": level e/ou dict de fields
    if args:
        for a in args[:2]:
            if isinstance(a, (str, int)):
                lvl = a
            elif isinstance(a, dict):
                fields.update(a)
            elif a is not None:
                # fallback: trata como string e coloca como detalhe
                fields.setdefault("detail", str(a))

    lvl_value = _resolve_level(lvl)

    ftxt = fmt_kv(fields)
    if ftxt:
        logger.log(lvl_value, "%s | %s", msg, ftxt)
    else:
        logger.log(lvl_value, "%s", msg)


@contextmanager
def step(logger: Optional[logging.Logger], name: str, **fields: Any) -> Iterator[None]:
    """
    Context manager de steps.

    Mantém o padrão atual:
      [STEP] ...
      [STEP_OK] ... dt_ms=...
      [STEP_FAIL] ... dt_ms=... (com stacktrace)

    E, se existir, emite “legacy report” (sem timestamps) via soma_app.infra.report.
    """
    if logger is None:
        logger = logging.getLogger(__name__)

    # Report “legacy” (não quebra se o módulo não existir)
    _report = None
    try:
        from soma_app.infra import report as _report  # type: ignore
        _report.on_step_start(name, dict(fields))
    except Exception:
        _report = None

    ftxt = fmt_kv(fields)
    if ftxt:
        logger.info("[STEP] %s | %s", name, ftxt)
    else:
        logger.info("[STEP] %s", name)

    t0 = time.perf_counter()
    try:
        yield
    except Exception as e:
        dt_ms = int((time.perf_counter() - t0) * 1000)
        if ftxt:
            logger.error("[STEP_FAIL] %s | dt_ms=%s | %s", name, dt_ms, ftxt, exc_info=True)
        else:
            logger.error("[STEP_FAIL] %s | dt_ms=%s", name, dt_ms, exc_info=True)

        try:
            if _report is not None:
                _report.on_step_fail(name, dict(fields), dt_ms, type(e).__name__)
        except Exception:
            pass
        raise
    else:
        dt_ms = int((time.perf_counter() - t0) * 1000)
        if ftxt:
            logger.info("[STEP_OK] %s | dt_ms=%s | %s", name, dt_ms, ftxt)
        else:
            logger.info("[STEP_OK] %s | dt_ms=%s", name, dt_ms)

        try:
            if _report is not None:
                _report.on_step_ok(name, dict(fields), dt_ms)
        except Exception:
            pass


__all__ = ["step", "log_kv", "fmt_kv", "new_run_id"]