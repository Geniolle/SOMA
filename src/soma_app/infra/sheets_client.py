# src/soma_app/infra/sheets_client.py
from __future__ import annotations

import logging
import os
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import gspread

logger = logging.getLogger(__name__)


# Carrega .env automaticamente (best effort)
try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except Exception:
    pass


def _norm(s: Any) -> str:
    return str(s or "").strip().strip('"').strip("'")


def _get_attr(settings: Any, name: str) -> Optional[Any]:
    if settings is None:
        return None
    if isinstance(settings, dict):
        return settings.get(name)
    return getattr(settings, name, None)


def _first_value(*vals: Any) -> Optional[Any]:
    for v in vals:
        if v is None:
            continue
        if isinstance(v, str):
            if v.strip() == "":
                continue
            return v
        return v
    return None


def _resolve_credentials_path(settings: Any) -> str:
    env_path = _first_value(
        os.getenv("GOOGLE_CREDENTIALS_PATH"),
        os.getenv("GOOGLE_APPLICATION_CREDENTIALS"),
    )

    cand = _first_value(
        env_path,
        _get_attr(settings, "google_credentials_path"),
        _get_attr(settings, "GOOGLE_CREDENTIALS_PATH"),
        _get_attr(settings, "credentials_path"),
        _get_attr(settings, "CREDENTIALS_PATH"),
        _get_attr(settings, "service_account_file"),
        _get_attr(settings, "SERVICE_ACCOUNT_FILE"),
        _get_attr(settings, "gspread_credentials_path"),
    )

    path = _norm(cand) if cand else ""
    if not path:
        raise AttributeError(
            "Não encontrei o caminho das credenciais Google. "
            "Define GOOGLE_CREDENTIALS_PATH (env/.env) ou adiciona google_credentials_path no Settings."
        )

    path = os.path.expandvars(path)
    path = os.path.expanduser(path)

    # relativo -> absoluto
    if not os.path.isabs(path):
        path = os.path.abspath(path)

    # fallback: se o path não existe mas o ficheiro existe na raiz do projeto (cwd)
    if not os.path.exists(path):
        base = os.path.basename(path)
        alt = os.path.abspath(os.path.join(os.getcwd(), base))
        if os.path.exists(alt):
            path = alt

    if not os.path.exists(path):
        raise FileNotFoundError(f"Ficheiro de credenciais não encontrado: {path}")

    return path


def _extract_sheet_id_from_url(url: str) -> Optional[str]:
    # https://docs.google.com/spreadsheets/d/<ID>/edit...
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    return m.group(1) if m else None


def _resolve_spreadsheet(settings: Any) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Retorna (spreadsheet_id, spreadsheet_name, spreadsheet_url)
    """
    sid = _first_value(
        os.getenv("SPREADSHEET_ID"),
        _get_attr(settings, "spreadsheet_id"),
        _get_attr(settings, "SPREADSHEET_ID"),
    )

    surl = _first_value(
        os.getenv("SPREADSHEET_URL"),
        _get_attr(settings, "spreadsheet_url"),
        _get_attr(settings, "SPREADSHEET_URL"),
    )

    sname = _first_value(
        os.getenv("SPREADSHEET_NAME"),
        os.getenv("SPREADSHEET"),
        _get_attr(settings, "spreadsheet_name"),
        _get_attr(settings, "SPREADSHEET_NAME"),
        _get_attr(settings, "spreadsheet"),
        _get_attr(settings, "SPREADSHEET"),
    )

    sid = _norm(sid) if sid else ""
    sname = _norm(sname) if sname else ""
    surl = _norm(surl) if surl else ""

    return (sid or None, sname or None, surl or None)


class SheetsClient:
    """
    Cliente Google Sheets (gspread) resiliente a Settings como objeto ou módulo.

    Suporta:
      - GOOGLE_CREDENTIALS_PATH / GOOGLE_APPLICATION_CREDENTIALS
      - SPREADSHEET_ID / SPREADSHEET_NAME / SPREADSHEET_URL
    """

    def __init__(self, settings: Any):
        self.settings = settings

        creds_path = _resolve_credentials_path(settings)
        self._gc = gspread.service_account(filename=str(creds_path))

        spreadsheet_id, spreadsheet_name, spreadsheet_url = _resolve_spreadsheet(settings)

        if spreadsheet_url:
            # preferir open_by_url; fallback para open_by_key
            try:
                self._sh = self._gc.open_by_url(spreadsheet_url)
            except Exception:
                sid = _extract_sheet_id_from_url(spreadsheet_url)
                if not sid:
                    raise AttributeError(f"Não consegui extrair SPREADSHEET_ID da URL: {spreadsheet_url}")
                self._sh = self._gc.open_by_key(sid)

        elif spreadsheet_id:
            self._sh = self._gc.open_by_key(spreadsheet_id)

        elif spreadsheet_name:
            self._sh = self._gc.open(spreadsheet_name)

        else:
            raise AttributeError(
                "Não encontrei SPREADSHEET_URL, SPREADSHEET_ID nem SPREADSHEET_NAME. "
                "Define no .env/env ou no Settings."
            )

        try:
            title = self._sh.title
        except Exception:
            title = spreadsheet_name or spreadsheet_id or "-"
        logger.info("SheetsClient OK | spreadsheet=%s", title)

    def _ws(self, ws: str):
        return self._sh.worksheet(ws)

    # -------- leituras --------
    def get_all_records(self, ws: str) -> List[Dict[str, Any]]:
        return self._ws(ws).get_all_records()

    def read_all_records(self, ws: str) -> List[Dict[str, Any]]:
        return self.get_all_records(ws)

    def read_records(self, ws: str) -> List[Dict[str, Any]]:
        return self.get_all_records(ws)

    def get_header(self, ws: str, row: int = 1) -> List[str]:
        values = self._ws(ws).row_values(row)
        return [str(x).strip() for x in values]

    def get_row(self, ws: str, row: int) -> List[str]:
        return self._ws(ws).row_values(row)

    # -------- updates --------
    def update_cell(self, ws: str, row: int, col: int, value: Any) -> None:
        self._ws(ws).update_cell(row, col, value)

    def batch_update(self, ws: str, ranges: Sequence[Tuple[str, List[List[Any]]]]) -> None:
        if not ranges:
            return

        data = []
        for a1, values in ranges:
            a1 = str(a1)
            rng = a1 if "!" in a1 else f"{ws}!{a1}"
            data.append({"range": rng, "values": values})

        try:
            self._sh.values_batch_update({"valueInputOption": "USER_ENTERED", "data": data})
        except Exception:
            w = self._ws(ws)
            for item in data:
                rng = item["range"].split("!", 1)[1]
                w.update(rng, item["values"], value_input_option="USER_ENTERED")

        logger.info("Batch update OK | ws=%s | ranges=%s", ws, len(ranges))

    def batch_update_rows(self, ws: str, updates: List[Dict[str, Any]]) -> None:
        if not updates:
            return

        header = self.get_header(ws, row=1)
        idx = {h: i + 1 for i, h in enumerate(header)}  # 1-based

        ranges: List[Tuple[str, List[List[Any]]]] = []
        for upd in updates:
            row = int(upd["row"])
            for col_name, val in upd.items():
                if col_name == "row":
                    continue
                if col_name not in idx:
                    continue
                col = idx[col_name]
                a1 = f"{_col_letter(col)}{row}"
                ranges.append((a1, [[val]]))

        if ranges:
            self.batch_update(ws, ranges)


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s