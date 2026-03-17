from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)


DOC_COL_DEFAULT = "DOC. SOMA"
TIPO_COL_DEFAULT = "TIPO"
STATUS_COL_DEFAULT = "STATUS"

LOCK_VALUE_DEFAULT = "Em processamento"


def _now_ts_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _norm(s: Any) -> str:
    return str(s or "").strip()


def _norm_tipo(v: Any) -> str:
    t = _norm(v).lower()
    if t in {"entrada", "entradas"}:
        return "Entrada"
    if t in {"saída", "saida", "saidas", "saídas"}:
        return "Saída"
    # ✅ NOVO: Transferência
    if t in {"transferência", "transferencia", "transferências", "transferencias"}:
        return "Transferência"
    return _norm(v)


def _bool_env(name: str, default: bool = False) -> bool:
    v = (os.getenv(name) or "").strip().lower()
    if v in {"1", "true", "yes", "y", "on"}:
        return True
    if v in {"0", "false", "no", "n", "off"}:
        return False
    return default


def _int_env(name: str, default: int) -> int:
    v = (os.getenv(name) or "").strip()
    try:
        return int(v)
    except Exception:
        return default


def _col_letter(n: int) -> str:
    # 1 -> A
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _a1(col_idx_1based: int, row_idx_1based: int) -> str:
    return f"{_col_letter(col_idx_1based)}{row_idx_1based}"


class SheetsTable:
    """
    Adapter resiliente para o SheetsClient.
    - lê header + records
    - aplica updates por col_name/row_index
    """

    def __init__(self, sheets: Any, ws: str):
        self.sheets = sheets
        self.ws = ws
        self.header: List[str] = []
        self.records: List[Dict[str, Any]] = []  # sem row index
        self._header_map: Dict[str, int] = {}  # col_name -> 1-based idx

    def load(self) -> None:
        self.header = self._read_header()
        self._header_map = {h: i + 1 for i, h in enumerate(self.header)}
        self.records = self._read_records()

    def col_idx(self, col_name: str) -> Optional[int]:
        return self._header_map.get(col_name)

    def has_col(self, col_name: str) -> bool:
        return self.col_idx(col_name) is not None

    def get_records_with_row(self) -> List[Dict[str, Any]]:
        # records começam na linha 2 (header é linha 1)
        out: List[Dict[str, Any]] = []
        for i, r in enumerate(self.records, start=2):
            rr = dict(r)
            rr["row"] = i
            out.append(rr)
        return out

    def batch_update_cells(self, updates: Sequence[Tuple[int, str, Any]]) -> None:
        """
        updates: lista de (row_idx_1based, col_name, value)
        """
        if not updates:
            return

        # tenta métodos “altos” do SheetsClient (se existirem)
        for meth in ("batch_update_rows", "batch_update_row_values", "update_rows"):
            fn = getattr(self.sheets, meth, None)
            if callable(fn):
                # monta payload row-based
                payload = []
                for row_idx, col_name, value in updates:
                    payload.append({"row": row_idx, col_name: value})
                try:
                    fn(ws=self.ws, updates=payload)  # padrão comum
                    return
                except TypeError:
                    fn(self.ws, payload)
                    return

        # tenta batch_update por ranges (A1)
        fn = getattr(self.sheets, "batch_update", None)
        if callable(fn):
            ranges_payload = []
            for row_idx, col_name, value in updates:
                col_idx = self.col_idx(col_name)
                if not col_idx:
                    raise RuntimeError(f"Coluna não encontrada no header: {col_name}")
                cell = _a1(col_idx, row_idx)
                # alguns wrappers querem "A1", outros "WS!A1"
                ranges_payload.append((f"{self.ws}!{cell}", [[value]]))

            # tenta várias assinaturas
            try:
                fn(ws=self.ws, ranges=ranges_payload)
                return
            except TypeError:
                try:
                    fn(self.ws, ranges_payload)
                    return
                except TypeError:
                    pass

        # tenta update_cell (singular)
        fn = getattr(self.sheets, "update_cell", None)
        if callable(fn):
            for row_idx, col_name, value in updates:
                col_idx = self.col_idx(col_name)
                if not col_idx:
                    raise RuntimeError(f"Coluna não encontrada no header: {col_name}")
                try:
                    fn(ws=self.ws, row=row_idx, col=col_idx, value=value)
                except TypeError:
                    fn(self.ws, row_idx, col_idx, value)
            return

        raise RuntimeError(
            "SheetsClient: não encontrei método suportado para update (batch_update/update_cell/update_rows). "
            "Ajusta o SheetsTable.batch_update_cells() para a API real do teu SheetsClient."
        )

    def _read_header(self) -> List[str]:
        # tenta métodos comuns
        for meth in ("get_header", "read_header", "header", "get_row", "read_row", "row_values"):
            fn = getattr(self.sheets, meth, None)
            if callable(fn):
                try:
                    h = fn(ws=self.ws, row=1)
                except TypeError:
                    try:
                        h = fn(self.ws, 1)
                    except TypeError:
                        h = fn(self.ws)
                if isinstance(h, dict) and "values" in h:
                    h = h["values"]
                if isinstance(h, (list, tuple)):
                    return [str(x).strip() for x in h if str(x).strip()]

        # fallback: tenta ler range 1:1
        fn = getattr(self.sheets, "get_values", None)
        if callable(fn):
            try:
                vals = fn(ws=self.ws, a1=f"{self.ws}!1:1")
            except TypeError:
                vals = fn(self.ws, f"{self.ws}!1:1")
            if vals and isinstance(vals, list):
                row = vals[0] if vals else []
                return [str(x).strip() for x in row if str(x).strip()]

        raise RuntimeError("Não consegui ler o header da worksheet via SheetsClient.")

    def _read_records(self) -> List[Dict[str, Any]]:
        # tenta métodos comuns que retornam list[dict]
        for meth in ("get_all_records", "read_all_records", "read_records", "get_records"):
            fn = getattr(self.sheets, meth, None)
            if callable(fn):
                try:
                    recs = fn(ws=self.ws)
                except TypeError:
                    recs = fn(self.ws)
                if isinstance(recs, list):
                    return [dict(r) for r in recs]

        # fallback: tenta read_table com header
        fn = getattr(self.sheets, "read_table", None)
        if callable(fn):
            try:
                out = fn(ws=self.ws)
            except TypeError:
                out = fn(self.ws)
            if isinstance(out, dict) and "records" in out:
                return [dict(r) for r in out["records"]]

        raise RuntimeError("Não consegui ler records da worksheet via SheetsClient.")


@dataclass
class PreprocessResult:
    total: int
    parsed: int
    parse_errors: int
    inflight: int
    batch_size: int
    capacity_new: int
    workset: List[Dict[str, Any]]
    newly_locked: int
    invalid: int


def preprocess_contaordem(
    sheets: Any,
    *,
    ws: str,
    run_id: str,
    batch: int,
    batch_size: Optional[int] = None,
    doc_col: str = DOC_COL_DEFAULT,
    tipo_col: str = TIPO_COL_DEFAULT,
    status_col: str = STATUS_COL_DEFAULT,
    lock_value: str = LOCK_VALUE_DEFAULT,
) -> PreprocessResult:
    """
    Preprocess:
    - lê sheet
    - calcula inflight (linhas já lockadas)
    - escolhe workset por prioridade:
        Entrada primeiro; se houver Entrada pendente, NÃO locka Saídas/Transferências neste batch
        depois Saída
        depois Transferência
    - locka APENAS o workset (doc_col = 'Em processamento')
    """
    bsz = batch_size or _int_env("BATCH_SIZE", 20)

    t = SheetsTable(sheets, ws)
    t.load()

    rows = t.get_records_with_row()
    total = len(rows)
    parsed = 0
    parse_errors = 0
    inflight = 0
    invalid = 0

    candidates_entrada: List[Dict[str, Any]] = []
    candidates_saida: List[Dict[str, Any]] = []
    # ✅ NOVO: candidates transferência
    candidates_transfer: List[Dict[str, Any]] = []

    for r in rows:
        parsed += 1
        doc = _norm(r.get(doc_col))
        status = _norm(r.get(status_col)).upper()
        tipo = _norm_tipo(r.get(tipo_col))

        # inflight: já lockado
        if doc.lower() == lock_value.lower():
            inflight += 1
            continue

        # já processado (tem doc)
        if doc != "":
            continue

        # status ERRO/INVALIDO não entra
        if status in {"ERRO", "ERROR", "INVALIDO", "INVÁLIDO"}:
            continue

        # ✅ NOVO: aceita Transferência
        if tipo not in {"Entrada", "Saída", "Transferência"}:
            invalid += 1
            continue

        if tipo == "Entrada":
            candidates_entrada.append(r)
        elif tipo == "Saída":
            candidates_saida.append(r)
        else:
            candidates_transfer.append(r)

    capacity_new = max(bsz - inflight, 0)

    # prioridade: Entrada -> Saída -> Transferência
    if candidates_entrada:
        candidates_entrada.sort(key=lambda x: int(x.get("row", 0)))
        workset = candidates_entrada[:capacity_new]
    elif candidates_saida:
        candidates_saida.sort(key=lambda x: int(x.get("row", 0)))
        workset = candidates_saida[:capacity_new]
    else:
        candidates_transfer.sort(key=lambda x: int(x.get("row", 0)))
        workset = candidates_transfer[:capacity_new]

    # lock apenas o workset
    updates: List[Tuple[int, str, Any]] = []
    for r in workset:
        updates.append((int(r["row"]), doc_col, lock_value))

    if updates:
        t.batch_update_cells(updates)

    newly_locked = len(updates)

    logger.info(
        "PREPROCESS | total=%s | parsed=%s | parse_errors=%s | inflight=%s | batch_size=%s | capacity_new=%s | workset=%s | newly_locked=%s | invalid=%s",
        total, parsed, parse_errors, inflight, bsz, capacity_new, len(workset), newly_locked, invalid,
    )

    return PreprocessResult(
        total=total,
        parsed=parsed,
        parse_errors=parse_errors,
        inflight=inflight,
        batch_size=bsz,
        capacity_new=capacity_new,
        workset=workset,
        newly_locked=newly_locked,
        invalid=invalid,
    )