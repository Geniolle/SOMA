# src/soma_app/infra/audit.py
from __future__ import annotations

import contextvars
import json
import logging
import os
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple


_AUDIT_CTX: contextvars.ContextVar[Dict[str, Any]] = contextvars.ContextVar("AUDIT_CTX", default={})


def _now_ymd() -> str:
    return datetime.now().strftime("%Y%m%d")


def _safe_str(v: Any) -> str:
    if v is None:
        return "-"
    if isinstance(v, (int, float, bool)):
        return str(v)
    if isinstance(v, (dict, list, tuple)):
        return json.dumps(v, ensure_ascii=False, separators=(",", ":"))
    s = str(v)
    # Se tiver espaços ou separadores comuns de log, serializa com aspas
    if any(c in s for c in [" ", "|", "\t", "\n", "\r"]):
        return json.dumps(s, ensure_ascii=False)
    return s


def _kv(**fields: Any) -> str:
    parts: List[str] = []
    for k, v in fields.items():
        if v is None:
            continue
        parts.append(f"{k}={_safe_str(v)}")
    return " ".join(parts)


@dataclass
class StepResult:
    step: str
    dt_ms: int
    status: str  # OK / ERR
    extra: Dict[str, Any]


def _get_ctx() -> Dict[str, Any]:
    return dict(_AUDIT_CTX.get() or {})


def _set_ctx(**updates: Any) -> None:
    ctx = _get_ctx()
    ctx.update({k: v for k, v in updates.items() if v is not None})
    _AUDIT_CTX.set(ctx)


def _clear_ctx() -> None:
    _AUDIT_CTX.set({})


def get_audit_logger() -> logging.Logger:
    return logging.getLogger("soma_audit")


def audit_event(event: str, **fields: Any) -> None:
    """
    Emite uma linha de auditoria com contexto + campos.
    """
    logger = get_audit_logger()
    ctx = _get_ctx()

    # Merge: ctx primeiro (fixos), depois fields (override)
    merged: Dict[str, Any] = {}
    merged.update(ctx)
    merged.update(fields)

    msg = f"event={event} " + _kv(**merged)
    logger.info(msg)


@contextmanager
def audit_row(
    *,
    run_id: str,
    batch: int,
    row: int,
    tipo: str,
    ws: Optional[str] = None,
    payload: Optional[Dict[str, Any]] = None,
) -> Any:
    """
    Contexto de auditoria por linha.
    Gera ROW_START e ROW_END (com resumo de steps).
    """
    _set_ctx(run_id=run_id, batch=batch, row=row, tipo=tipo, ws=ws)
    ctx = _get_ctx()
    ctx["__t0__"] = time.perf_counter()
    ctx["__steps__"] = []  # type: ignore
    _AUDIT_CTX.set(ctx)

    if payload:
        # Evita gravar segredos: filtra aqui se necessário
        audit_event("ROW_START", **payload)
    else:
        audit_event("ROW_START")

    status = "OK"
    doc = None
    try:
        yield
    except Exception as e:
        status = "ERR"
        # Mensagem curta para auditoria (stacktrace fica no APP log)
        audit_event("ROW_ERR", err=type(e).__name__, msg=str(e)[:180])
        raise
    finally:
        total_ms = int((time.perf_counter() - _get_ctx().get("__t0__", time.perf_counter())) * 1000)
        steps: List[StepResult] = _get_ctx().get("__steps__", [])  # type: ignore

        # Tenta inferir doc se algum step trouxe
        for s in steps:
            if "doc" in s.extra and s.extra.get("doc"):
                doc = s.extra.get("doc")

        steps_txt = "; ".join([f"{s.step}:{s.dt_ms}ms {s.status}" for s in steps])
        audit_event("ROW_END", status=status, doc=doc, total_ms=total_ms, steps=steps_txt)

        _clear_ctx()


@contextmanager
def audit_step(step: str, **fields: Any) -> Any:
    """
    Mede tempo e registra STEP_OK/STEP_ERR.
    O "início" fica em DEBUG no audit (não polui).
    """
    logger = get_audit_logger()
    t0 = time.perf_counter()

    # start em debug (opcional)
    logger.debug("event=STEP " + _kv(**_get_ctx(), step=step, **fields))

    try:
        yield
    except Exception as e:
        dt_ms = int((time.perf_counter() - t0) * 1000)
        extra = dict(fields)
        extra.update({"err": type(e).__name__, "msg": str(e)[:180]})

        _append_step(step=step, dt_ms=dt_ms, status="ERR", extra=extra)
        audit_event("STEP_ERR", step=step, dt_ms=dt_ms, **extra)
        raise
    else:
        dt_ms = int((time.perf_counter() - t0) * 1000)
        extra = dict(fields)

        _append_step(step=step, dt_ms=dt_ms, status="OK", extra=extra)
        audit_event("STEP_OK", step=step, dt_ms=dt_ms, **extra)


def _append_step(*, step: str, dt_ms: int, status: str, extra: Dict[str, Any]) -> None:
    ctx = _get_ctx()
    steps: List[StepResult] = ctx.get("__steps__", [])  # type: ignore
    steps.append(StepResult(step=step, dt_ms=dt_ms, status=status, extra=extra))
    ctx["__steps__"] = steps
    _AUDIT_CTX.set(ctx)