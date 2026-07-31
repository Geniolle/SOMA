from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from soma_app.infra.sheets_client import SheetsClient


CONTAORDEM_SHEET = "CONTAORDEM"
T_EXTRATO_SHEET = "T_EXTRATO"
VALIDATION_COL = "VALIDADO_EXTRATO"

PROCESSO_COL = "PROCESSO"
DOC_COL = "DOC. SOMA"
DATA_COL = "DATA MOV."
DESC_COL = "DESCRIÇÃO"
AMOUNT_COL = "IMPORTÂNCIA"
ID_INTERNO_COL = "ID_INTERNO"
EM_PROCESSAMENTO = "Em processamento"


@dataclass(frozen=True)
class SheetsSettings:
    google_credentials_path: Path
    spreadsheet_url: str
    spreadsheet_id: str = ""
    spreadsheet_name: str = ""


def _load_dotenv_file(path: Path) -> None:
    if load_dotenv is None:
        return
    if path.exists():
        load_dotenv(dotenv_path=str(path), override=False)


def _load_settings() -> Settings:
    _load_dotenv_file(ROOT / "deploy" / ".env")
    _load_dotenv_file(ROOT / ".env")

    env_file = os.getenv("ENV_FILE")
    if env_file:
        _load_dotenv_file(Path(env_file))

    cred_value = (os.getenv("GOOGLE_CREDENTIALS_PATH") or "").strip()
    spreadsheet_url = (os.getenv("SPREADSHEET_URL") or "").strip()
    spreadsheet_id = (os.getenv("SPREADSHEET_ID") or "").strip()
    spreadsheet_name = (os.getenv("SPREADSHEET_NAME") or os.getenv("SPREADSHEET") or "").strip()

    if not cred_value:
        raise ValueError("GOOGLE_CREDENTIALS_PATH em falta para aceder ao Google Sheets.")
    if not spreadsheet_url and not spreadsheet_id and not spreadsheet_name:
        raise ValueError("SPREADSHEET_URL, SPREADSHEET_ID ou SPREADSHEET_NAME em falta.")

    cred_path = Path(cred_value).expanduser()
    if not cred_path.is_absolute():
        candidate = (ROOT / cred_path).resolve()
        cred_path = candidate if candidate.exists() else cred_path.resolve()

    return SheetsSettings(
        google_credentials_path=cred_path,
        spreadsheet_url=spreadsheet_url,
        spreadsheet_id=spreadsheet_id,
        spreadsheet_name=spreadsheet_name,
    )


def _strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_text(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = _strip_accents(text).upper()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def normalize_date(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")
    if isinstance(value, (int, float)):
        try:
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=float(value))).strftime("%Y%m%d")
        except Exception:
            return str(value).strip()

    text = str(value).strip()
    if not text:
        return ""

    for fmt in (
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%Y%m%d",
    ):
        try:
            return datetime.strptime(text, fmt).strftime("%Y%m%d")
        except Exception:
            pass
    return normalize_text(text)


def normalize_amount(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return f"{abs(float(value)):.2f}"

    text = str(value).strip()
    if not text:
        return ""

    text = text.replace(" ", "")
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")

    try:
        return f"{abs(float(text)):.2f}"
    except Exception:
        return normalize_text(text)


def build_match_key(date_value: Any, desc_value: Any, amount_value: Any) -> Tuple[str, str, str]:
    return (
        normalize_date(date_value),
        normalize_text(desc_value),
        normalize_amount(amount_value),
    )


def row_match_key(row: Dict[str, Any]) -> Tuple[str, str, str]:
    return build_match_key(row.get(DATA_COL), row.get(DESC_COL), row.get(AMOUNT_COL))


def _process_state(row: Dict[str, Any]) -> str:
    doc = str(row.get(DOC_COL, "")).strip()
    processo = str(row.get(PROCESSO_COL, "")).strip()
    if not doc:
        return "doc_vazio"
    if doc.lower() == EM_PROCESSAMENTO.lower():
        return "em_processamento"
    if processo != T_EXTRATO_SHEET:
        return "outro_processo"
    if str(row.get(VALIDATION_COL, "")).strip():
        return "ja_validado"
    return "elegivel"


def _is_processable(row: Dict[str, Any]) -> bool:
    return _process_state(row) == "elegivel"


@dataclass(frozen=True)
class MatchResult:
    contaordem_row: int
    t_extrato_row: int
    contaordem: Dict[str, Any]
    origem: Dict[str, Any]
    key: Tuple[str, str, str]
    id_interno_equal: bool
    doc_equal: bool


@dataclass
class ValidationReport:
    contaordem_lidas: int = 0
    t_extrato_lidas: int = 0
    elegiveis: int = 0
    doc_vazio: int = 0
    em_processamento: int = 0
    outro_processo: int = 0
    ja_validado: int = 0
    matches: int = 0
    contaordem_linhas_alteradas: int = 0
    t_extrato_linhas_alteradas: int = 0
    contaordem_id_interno_alterado: int = 0
    t_extrato_doc_soma_alterado: int = 0
    sem_match_restante: int = 0


def ensure_validation_column(client: SheetsClient, sheet_name: str, column_name: str) -> int:
    header = client.get_header(sheet_name, row=1)
    if column_name in header:
        return header.index(column_name) + 1

    col_idx = len(header) + 1
    client.update_cell(sheet_name, 1, col_idx, column_name)
    return col_idx


def _read_records(client: SheetsClient, sheet_name: str) -> List[Dict[str, Any]]:
    getter = getattr(client, "get_all_records_raw", None)
    if callable(getter):
        return getter(sheet_name)
    return client.get_all_records(sheet_name)


def _row_number_for_record(index: int) -> int:
    return index + 2


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _header_map(header: List[str]) -> Dict[str, int]:
    return {str(name).strip(): idx + 1 for idx, name in enumerate(header) if str(name).strip()}


def _build_index(records: Iterable[Dict[str, Any]]) -> Dict[Tuple[str, str, str], List[Tuple[int, Dict[str, Any]]]]:
    index: Dict[Tuple[str, str, str], List[Tuple[int, Dict[str, Any]]]] = {}
    for i, row in enumerate(records):
        key = row_match_key(row)
        if not all(key):
            continue
        index.setdefault(key, []).append((_row_number_for_record(i), row))
    return index


def summarize_rows(contaordem_rows: List[Dict[str, Any]]) -> ValidationReport:
    report = ValidationReport(contaordem_lidas=len(contaordem_rows))
    for row in contaordem_rows:
        state = _process_state(row)
        if state == "doc_vazio":
            report.doc_vazio += 1
        elif state == "em_processamento":
            report.em_processamento += 1
        elif state == "outro_processo":
            report.outro_processo += 1
        elif state == "ja_validado":
            report.ja_validado += 1
        elif state == "elegivel":
            report.elegiveis += 1
    return report


def find_single_match(
    contaordem_rows: List[Dict[str, Any]],
    t_extrato_rows: List[Dict[str, Any]],
) -> Optional[MatchResult]:
    indexed = _build_index(t_extrato_rows)

    for i, row in enumerate(contaordem_rows):
        if not _is_processable(row):
            continue

        key = row_match_key(row)
        candidates = indexed.get(key, [])
        if not candidates:
            continue

        contaordem_row = _row_number_for_record(i)
        origem_row, origem = candidates[0]
        contaordem_doc = str(row.get(DOC_COL, "")).strip()
        origem_doc = str(origem.get(DOC_COL, "")).strip()
        contaordem_id = str(row.get(ID_INTERNO_COL, "")).strip()
        origem_id = str(origem.get(ID_INTERNO_COL, "")).strip()
        return MatchResult(
            contaordem_row=contaordem_row,
            t_extrato_row=origem_row,
            contaordem=row,
            origem=origem,
            key=key,
            id_interno_equal=(contaordem_id == origem_id),
            doc_equal=(contaordem_doc == origem_doc),
        )

    return None


def sync_match(
    client: SheetsClient,
    match: MatchResult,
    contaordem_id_col: int,
    validation_col: int,
    origem_doc_col: int,
) -> Tuple[List[Tuple[str, List[List[Any]]]], List[Tuple[str, List[List[Any]]]], bool, bool]:
    origem_id = str(match.origem.get(ID_INTERNO_COL, "")).strip()
    contaordem_id = str(match.contaordem.get(ID_INTERNO_COL, "")).strip()
    origem_doc = str(match.origem.get(DOC_COL, "")).strip()
    contaordem_doc = str(match.contaordem.get(DOC_COL, "")).strip()

    contaordem_changed = False
    origem_changed = False

    if origem_id and origem_id != contaordem_id:
        contaordem_changed = True

    if contaordem_doc and contaordem_doc != origem_doc:
        origem_changed = True

    contaordem_payload: List[Tuple[str, List[List[Any]]]] = []
    if origem_id and origem_id != contaordem_id:
        contaordem_payload.append((f"{_col_letter(contaordem_id_col)}{match.contaordem_row}", [[origem_id]]))
    contaordem_payload.append((f"{_col_letter(validation_col)}{match.contaordem_row}", [["VALIDADO"]]))

    origem_payload: List[Tuple[str, List[List[Any]]]] = []
    if contaordem_doc and contaordem_doc != origem_doc:
        origem_payload.append((f"{_col_letter(origem_doc_col)}{match.t_extrato_row}", [[contaordem_doc]]))

    return contaordem_payload, origem_payload, contaordem_changed, origem_changed


def run(limit: Optional[int] = None, dry_run: bool = False, progress_every: int = 25) -> ValidationReport:
    settings = _load_settings()
    client = SheetsClient(settings)
    contaordem_header = client.get_header(CONTAORDEM_SHEET, row=1)
    t_extrato_header = client.get_header(T_EXTRATO_SHEET, row=1)

    contaordem_validation_col = ensure_validation_column(client, CONTAORDEM_SHEET, VALIDATION_COL)
    if VALIDATION_COL not in contaordem_header:
        contaordem_header = client.get_header(CONTAORDEM_SHEET, row=1)
    contaordem_header_idx = _header_map(contaordem_header)
    t_extrato_header_idx = _header_map(t_extrato_header)

    contaordem_rows = [dict(row) for row in _read_records(client, CONTAORDEM_SHEET)]
    t_extrato_rows = [dict(row) for row in _read_records(client, T_EXTRATO_SHEET)]

    report = summarize_rows(contaordem_rows)
    report.t_extrato_lidas = len(t_extrato_rows)

    processed = 0
    pending_contaordem_updates: List[Tuple[str, List[List[Any]]]] = []
    pending_t_extrato_updates: List[Tuple[str, List[List[Any]]]] = []
    matches_since_flush = 0

    def flush_pending() -> None:
        nonlocal pending_contaordem_updates, pending_t_extrato_updates, matches_since_flush
        if pending_contaordem_updates:
            client.batch_update(CONTAORDEM_SHEET, pending_contaordem_updates)
            pending_contaordem_updates = []
        if pending_t_extrato_updates:
            client.batch_update(T_EXTRATO_SHEET, pending_t_extrato_updates)
            pending_t_extrato_updates = []
        matches_since_flush = 0

    while limit is None or processed < limit:
        match = find_single_match(contaordem_rows, t_extrato_rows)
        if match is None:
            break

        if processed == 0 or (progress_every > 0 and processed % progress_every == 0):
            print(
                f"Progresso | match #{processed + 1} | "
                f"CONTAORDEM={match.contaordem_row} | T_EXTRATO={match.t_extrato_row} | key={match.key}"
            )

        if dry_run:
            print("Dry-run ativo: nenhuma escrita foi feita.")
            break

        contaordem_payload, origem_payload, contaordem_changed, origem_changed = sync_match(
            client,
            match,
            contaordem_header_idx[ID_INTERNO_COL],
            contaordem_validation_col,
            t_extrato_header_idx[DOC_COL],
        )

        contaordem_rows[match.contaordem_row - 2][VALIDATION_COL] = "VALIDADO"
        if contaordem_changed:
            contaordem_rows[match.contaordem_row - 2][ID_INTERNO_COL] = str(match.origem.get(ID_INTERNO_COL, "")).strip()
        if origem_changed:
            t_extrato_rows[match.t_extrato_row - 2][DOC_COL] = str(match.contaordem.get(DOC_COL, "")).strip()

        pending_contaordem_updates.extend(contaordem_payload)
        pending_t_extrato_updates.extend(origem_payload)
        matches_since_flush += 1

        if matches_since_flush >= 25 or len(pending_contaordem_updates) >= 60 or len(pending_t_extrato_updates) >= 60:
            flush_pending()

        report.matches += 1
        report.contaordem_linhas_alteradas += int(contaordem_changed)
        report.t_extrato_linhas_alteradas += int(origem_changed)
        report.contaordem_id_interno_alterado += int(
            str(match.origem.get(ID_INTERNO_COL, "")).strip() != str(match.contaordem.get(ID_INTERNO_COL, "")).strip()
        )
        report.t_extrato_doc_soma_alterado += int(
            str(match.contaordem.get(DOC_COL, "")).strip() != str(match.origem.get(DOC_COL, "")).strip()
        )

        processed += 1
    flush_pending()

    remaining = 0
    for row in contaordem_rows:
        if _is_processable(row):
            remaining += 1
    report.sem_match_restante = remaining
    return report


def print_report(report: ValidationReport) -> None:
    total_nao_habilitados = report.doc_vazio + report.em_processamento + report.outro_processo + report.ja_validado
    print("\nResumo final")
    print(f"- CONTAORDEM lidas: {report.contaordem_lidas}")
    print(f"- T_EXTRATO lidas: {report.t_extrato_lidas}")
    print(f"- Elegíveis para validação: {report.elegiveis}")
    print(f"- Não habilitados por DOC. SOMA vazio: {report.doc_vazio}")
    print(f"- Não habilitados por DOC. SOMA = Em processamento: {report.em_processamento}")
    print(f"- Não habilitados por outro PROCESSO: {report.outro_processo}")
    print(f"- Já validados e ignorados: {report.ja_validado}")
    print(f"- Total não habilitados a validação: {total_nao_habilitados}")
    print(f"- Registos validados com match: {report.matches}")
    print(f"- Linhas de CONTAORDEM alteradas: {report.contaordem_linhas_alteradas}")
    print(f"- Linhas de T_EXTRATO alteradas: {report.t_extrato_linhas_alteradas}")
    print(f"- Alterações de ID_INTERNO na CONTAORDEM: {report.contaordem_id_interno_alterado}")
    print(f"- Alterações de DOC. SOMA na T_EXTRATO: {report.t_extrato_doc_soma_alterado}")
    print(f"- Elegíveis restantes sem match: {report.sem_match_restante}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validação completa de T_EXTRATO contra CONTAORDEM.")
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Limita a quantidade de registos validados nesta execução. Por omissão, processa tudo.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Mostra os matches sem escrever na sheet.")
    parser.add_argument(
        "--progress-every",
        type=int,
        default=25,
        help="Mostra progresso a cada N registos validados. Use 0 para silenciar.",
    )
    args = parser.parse_args()

    if args.limit is not None and args.limit < 1:
        raise ValueError("--limit deve ser >= 1")
    if args.progress_every < 0:
        raise ValueError("--progress-every deve ser >= 0")

    report = run(limit=args.limit, dry_run=args.dry_run, progress_every=args.progress_every)
    print_report(report)


if __name__ == "__main__":
    main()
