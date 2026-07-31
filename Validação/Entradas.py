from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata
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
SOURCE_SHEET = "DÍZIMOS/OFERTAS"
PROCESS_NAME = "DÍZIMOS/OFERTAS"
VALIDATION_COL = "VALIDADO_DIZIMOS_OFERTAS"

PROCESSO_COL = "PROCESSO"
DOC_COL = "DOC. SOMA"
DATA_COL = "DATA MOV."
AMOUNT_COL = "IMPORTÂNCIA"
EM_PROCESSAMENTO = "Em processamento"


@dataclass(frozen=True)
class SheetsSettings:
    google_credentials_path: Path
    spreadsheet_url: str
    spreadsheet_id: str = ""
    spreadsheet_name: str = ""


@dataclass(frozen=True)
class MatchResult:
    contaordem_row: int
    source_row: int
    contaordem: Dict[str, Any]
    source: Dict[str, Any]
    key: Tuple[str, str]
    doc_equal: bool


@dataclass
class ValidationReport:
    contaordem_lidas: int = 0
    source_lidas: int = 0
    elegiveis: int = 0
    doc_vazio: int = 0
    em_processamento: int = 0
    outro_processo: int = 0
    ja_validado: int = 0
    matches: int = 0
    contaordem_linhas_alteradas: int = 0
    source_linhas_alteradas: int = 0
    source_doc_soma_alterado: int = 0
    sem_match_restante: int = 0


def _load_dotenv_file(path: Path) -> None:
    if load_dotenv is None:
        return
    if path.exists():
        load_dotenv(dotenv_path=str(path), override=False)


def _load_settings() -> SheetsSettings:
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

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%Y%m%d"):
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

    text = str(value).strip().replace(" ", "")
    if not text:
        return ""
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


def build_match_key(date_value: Any, amount_value: Any) -> Tuple[str, str]:
    return normalize_date(date_value), normalize_amount(amount_value)


def _read_records(client: SheetsClient, sheet_name: str) -> List[Dict[str, Any]]:
    getter = getattr(client, "get_all_records_raw", None)
    if callable(getter):
        return [dict(row) for row in getter(sheet_name)]
    return [dict(row) for row in client.get_all_records(sheet_name)]


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _header_map(header: List[str]) -> Dict[str, int]:
    return {str(name).strip(): idx + 1 for idx, name in enumerate(header) if str(name).strip()}


def _row_number_for_record(index: int) -> int:
    return index + 2


def _is_processable(row: Dict[str, Any], validation_col: str) -> bool:
    doc = str(row.get(DOC_COL, "")).strip()
    processo = str(row.get(PROCESSO_COL, "")).strip()
    if not doc:
        return False
    if doc.lower() == EM_PROCESSAMENTO.lower():
        return False
    if processo != PROCESS_NAME:
        return False
    if str(row.get(validation_col, "")).strip():
        return False
    return True


def _source_keys(row: Dict[str, Any]) -> Tuple[str, str, str]:
    date_key = next((k for k in row.keys() if not str(k).strip()), "")
    amount_key = next((k for k in row.keys() if normalize_text(k) == "VALOR"), "")
    doc_key = next((k for k in row.keys() if normalize_text(k) == "DOCSOMA"), "")
    return str(date_key), str(amount_key), str(doc_key)


def _build_index(records: Iterable[Dict[str, Any]], date_key: str, amount_key: str) -> Dict[Tuple[str, str], List[Tuple[int, Dict[str, Any]]]]:
    index: Dict[Tuple[str, str], List[Tuple[int, Dict[str, Any]]]] = {}
    for i, row in enumerate(records):
        key = build_match_key(row.get(date_key), row.get(amount_key))
        if not all(key):
            continue
        index.setdefault(key, []).append((_row_number_for_record(i), row))
    return index


def summarize_rows(contaordem_rows: List[Dict[str, Any]], validation_col: str) -> ValidationReport:
    report = ValidationReport(contaordem_lidas=len(contaordem_rows))
    for row in contaordem_rows:
        doc = str(row.get(DOC_COL, "")).strip()
        processo = str(row.get(PROCESSO_COL, "")).strip()
        if not doc:
            report.doc_vazio += 1
        elif doc.lower() == EM_PROCESSAMENTO.lower():
            report.em_processamento += 1
        elif processo != PROCESS_NAME:
            report.outro_processo += 1
        elif str(row.get(validation_col, "")).strip():
            report.ja_validado += 1
        else:
            report.elegiveis += 1
    return report


def find_single_match(
    contaordem_rows: List[Dict[str, Any]],
    source_rows: List[Dict[str, Any]],
    validation_col: str,
    source_date_key: str,
    source_amount_key: str,
) -> Optional[MatchResult]:
    indexed = _build_index(source_rows, source_date_key, source_amount_key)

    for i, row in enumerate(contaordem_rows):
        if not _is_processable(row, validation_col):
            continue

        key = build_match_key(row.get(DATA_COL), row.get(AMOUNT_COL))
        candidates = indexed.get(key, [])
        if not candidates:
            continue

        source_row, source = candidates[0]
        source_doc_key = next((k for k in source.keys() if normalize_text(k) == "DOCSOMA"), "")
        doc_equal = str(row.get(DOC_COL, "")).strip() == str(source.get(source_doc_key, "")).strip()

        return MatchResult(
            contaordem_row=_row_number_for_record(i),
            source_row=source_row,
            contaordem=row,
            source=source,
            key=key,
            doc_equal=doc_equal,
        )

    return None


def sync_match(
    client: SheetsClient,
    match: MatchResult,
    contaordem_validation_col: int,
    source_doc_col: int,
) -> Tuple[List[Tuple[str, List[List[Any]]]], List[Tuple[str, List[List[Any]]]], bool, bool]:
    contaordem_doc = str(match.contaordem.get(DOC_COL, "")).strip()
    source_doc_key = next((k for k in match.source.keys() if normalize_text(k) == "DOCSOMA"), "")
    source_doc = str(match.source.get(source_doc_key, "")).strip()

    contaordem_payload = [(f"{_col_letter(contaordem_validation_col)}{match.contaordem_row}", [["VALIDADO"]])]
    source_payload: List[Tuple[str, List[List[Any]]]] = []
    if contaordem_doc and contaordem_doc != source_doc:
        source_payload.append((f"{_col_letter(source_doc_col)}{match.source_row}", [[contaordem_doc]]))

    contaordem_changed = True
    source_changed = bool(source_payload)
    return contaordem_payload, source_payload, contaordem_changed, source_changed


def run(limit: Optional[int] = None, dry_run: bool = False, progress_every: int = 25) -> ValidationReport:
    settings = _load_settings()
    client = SheetsClient(settings)

    contaordem_header = client.get_header(CONTAORDEM_SHEET, row=1)
    source_header = client.get_header(SOURCE_SHEET, row=1)
    contaordem_header_map = _header_map(contaordem_header)
    source_header_map = _header_map(source_header)

    contaordem_validation_col = contaordem_header_map.get(VALIDATION_COL)
    if contaordem_validation_col is None:
        contaordem_validation_col = len(contaordem_header_map) + 1
        client.update_cell(CONTAORDEM_SHEET, 1, contaordem_validation_col, VALIDATION_COL)
        contaordem_header = client.get_header(CONTAORDEM_SHEET, row=1)
        contaordem_header_map = _header_map(contaordem_header)

    source_doc_col = source_header_map.get("DOC. SOMA")
    if source_doc_col is None:
        raise ValueError("Não encontrei a coluna DOC. SOMA na sheet de origem.")

    source_rows = _read_records(client, SOURCE_SHEET)
    source_date_key, source_amount_key, _ = _source_keys(source_rows[0] if source_rows else {})

    contaordem_rows = _read_records(client, CONTAORDEM_SHEET)
    report = summarize_rows(contaordem_rows, VALIDATION_COL)
    report.source_lidas = len(source_rows)

    processed = 0
    pending_contaordem_updates: List[Tuple[str, List[List[Any]]]] = []
    pending_source_updates: List[Tuple[str, List[List[Any]]]] = []
    matches_since_flush = 0

    def flush_pending() -> None:
        nonlocal pending_contaordem_updates, pending_source_updates, matches_since_flush
        if pending_contaordem_updates:
            client.batch_update(CONTAORDEM_SHEET, pending_contaordem_updates)
            pending_contaordem_updates = []
        if pending_source_updates:
            client.batch_update(SOURCE_SHEET, pending_source_updates)
            pending_source_updates = []
        matches_since_flush = 0

    while limit is None or processed < limit:
        match = find_single_match(contaordem_rows, source_rows, VALIDATION_COL, source_date_key, source_amount_key)
        if match is None:
            break

        if processed == 0 or (progress_every > 0 and processed % progress_every == 0):
            print(
                f"Progresso | match #{processed + 1} | "
                f"CONTAORDEM={match.contaordem_row} | ORIGEM={match.source_row} | key={match.key}"
            )

        if dry_run:
            print("Dry-run ativo: nenhuma escrita foi feita.")
            break

        contaordem_payload, source_payload, contaordem_changed, source_changed = sync_match(
            client,
            match,
            contaordem_validation_col,
            source_doc_col,
        )

        contaordem_rows[match.contaordem_row - 2][VALIDATION_COL] = "VALIDADO"
        if source_changed:
            source_doc_key = next((k for k in match.source.keys() if normalize_text(k) == "DOCSOMA"), "")
            if source_doc_key:
                source_rows[match.source_row - 2][source_doc_key] = str(match.contaordem.get(DOC_COL, "")).strip()

        pending_contaordem_updates.extend(contaordem_payload)
        pending_source_updates.extend(source_payload)
        matches_since_flush += 1

        if matches_since_flush >= 25 or len(pending_contaordem_updates) >= 60 or len(pending_source_updates) >= 60:
            flush_pending()

        report.matches += 1
        report.contaordem_linhas_alteradas += int(contaordem_changed)
        report.source_linhas_alteradas += int(source_changed)
        report.source_doc_soma_alterado += int(source_changed)

        processed += 1

    flush_pending()

    remaining = 0
    for row in contaordem_rows:
        if _is_processable(row, VALIDATION_COL):
            remaining += 1
    report.sem_match_restante = remaining
    return report


def print_report(report: ValidationReport) -> None:
    total_nao_habilitados = report.doc_vazio + report.em_processamento + report.outro_processo + report.ja_validado
    print("\nResumo final")
    print(f"- CONTAORDEM lidas: {report.contaordem_lidas}")
    print(f"- Origem lida: {report.source_lidas}")
    print(f"- Elegíveis para validação: {report.elegiveis}")
    print(f"- Não habilitados por DOC. SOMA vazio: {report.doc_vazio}")
    print(f"- Não habilitados por DOC. SOMA = Em processamento: {report.em_processamento}")
    print(f"- Não habilitados por outro PROCESSO: {report.outro_processo}")
    print(f"- Já validados e ignorados: {report.ja_validado}")
    print(f"- Total não habilitados a validação: {total_nao_habilitados}")
    print(f"- Registos validados com match: {report.matches}")
    print(f"- Linhas de CONTAORDEM alteradas: {report.contaordem_linhas_alteradas}")
    print(f"- Linhas de origem alteradas: {report.source_linhas_alteradas}")
    print(f"- Alterações de DOC. SOMA na origem: {report.source_doc_soma_alterado}")
    print(f"- Elegíveis restantes sem match: {report.sem_match_restante}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validação de DÍZIMOS/OFERTAS contra CONTAORDEM.")
    parser.add_argument("--limit", type=int, default=None, help="Limita a quantidade de registos validados.")
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
