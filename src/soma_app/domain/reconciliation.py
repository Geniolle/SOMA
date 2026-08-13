from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from enum import Enum
from typing import Any, Iterable, Mapping, Sequence


class ReconciliationCategory(str, Enum):
    OK = "OK"
    EXCLUIDO = "EXCLUIDO"
    NAO_ENCONTRADO_CONTAORDEM = "NAO_ENCONTRADO_CONTAORDEM"
    NAO_ENCONTRADO_SOMA = "NAO_ENCONTRADO_SOMA"
    DIVERGENCIA = "DIVERGENCIA"
    DUPLICADO = "DUPLICADO"
    ORIGEM_INVALIDA = "ORIGEM_INVALIDA"


class Severity(str, Enum):
    CRITICO = "CRITICO"
    MEDIO = "MEDIO"
    INFORMATIVO = "INFORMATIVO"


@dataclass(frozen=True)
class ReconciliationIssue:
    code: str
    message: str
    field: str = ""
    origin_value: str = ""
    target_value: str = ""


@dataclass
class ReconciliationItem:
    source_name: str
    source_row: int
    id_interno: str
    doc_soma: str = ""
    contaordem_rows: list[int] = field(default_factory=list)
    soma_rows: list[int] = field(default_factory=list)
    category: ReconciliationCategory = ReconciliationCategory.OK
    severity: Severity = Severity.INFORMATIVO
    issues: list[ReconciliationIssue] = field(default_factory=list)
    warnings: list[ReconciliationIssue] = field(default_factory=list)
    excluded_from_soma: bool = False

    @property
    def flow(self) -> str:
        origem = "✅"
        conta = "✅" if self.contaordem_rows else "❌"
        if self.excluded_from_soma:
            soma = "⏭"
        elif self.soma_rows:
            soma = "✅"
        elif self.contaordem_rows:
            soma = "❌"
        else:
            soma = "—"
        return f"{self.source_name} {origem} → CONTAORDEM {conta} → SOMA {soma}"

    @property
    def action(self) -> str:
        if self.category == ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM:
            return "Verificar a carga/orquestração da origem para a CONTAORDEM."
        if self.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA:
            return "Verificar o DOC. SOMA e o envio/processamento deste lançamento no SOMA."
        if self.category == ReconciliationCategory.DUPLICADO:
            return "Rever duplicidade de ID/DOC. SOMA antes de qualquer correção automática."
        if self.category == ReconciliationCategory.DIVERGENCIA:
            return "Rever os campos divergentes e a associação entre origem, CONTAORDEM e SOMA."
        if self.category == ReconciliationCategory.ORIGEM_INVALIDA:
            return "Corrigir o registo na origem antes de continuar a conciliação."
        if self.category == ReconciliationCategory.EXCLUIDO:
            return "Sem ação: lançamento identificado como exclusão intencional do SOMA."
        return "Sem ação."


@dataclass
class ReconciliationReport:
    source_name: str
    year: int
    month: int
    items: list[ReconciliationItem]
    source_rows_total: int
    source_rows_outside_period: int = 0
    source_rows_invalid_date: int = 0

    def category_counts(self) -> Counter[ReconciliationCategory]:
        return Counter(item.category for item in self.items)

    @property
    def warning_count(self) -> int:
        return sum(len(item.warnings) for item in self.items)

    @property
    def reconcilable_count(self) -> int:
        return sum(item.category != ReconciliationCategory.EXCLUIDO for item in self.items)

    @property
    def ok_count(self) -> int:
        return sum(item.category == ReconciliationCategory.OK for item in self.items)

    @property
    def reconciliation_rate(self) -> Decimal:
        denom = self.reconcilable_count
        if denom <= 0:
            return Decimal("100")
        return (Decimal(self.ok_count) * Decimal("100")) / Decimal(denom)


@dataclass(frozen=True)
class _RowRef:
    row: int
    data: Mapping[str, Any]


_DOC_RE = re.compile(r"^\d+$")
_PLACEHOLDER_DOCS = {
    "",
    "EMPROCESSAMENTO",
    "TRANSFERIDO",
    "TRANSFERIDA",
    "DUPLICADO",
    "EMERRO",
    "ERRO",
    "ENCERRADO",
}


def normalize_text(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.upper().split())


def normalize_compact(value: Any) -> str:
    return "".join(ch for ch in normalize_text(value) if ch.isalnum())


def normalize_doc(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    m = re.fullmatch(r"(\d+)\.0+", text)
    if m:
        return m.group(1)
    return normalize_text(text).replace(" ", "")


def is_real_doc(value: Any) -> bool:
    return bool(_DOC_RE.fullmatch(normalize_doc(value)))


def parse_amount(value: Any) -> Decimal | None:
    if value is None:
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool):
        return Decimal(int(value))
    if isinstance(value, (int, float)):
        return Decimal(str(value))

    text = str(value).strip()
    if not text:
        return None

    negative_by_parentheses = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = text.replace("€", "").replace("EUR", "").replace("eur", "")
    text = text.replace("\u00a0", "").replace(" ", "")
    text = re.sub(r"[^0-9,\.\-+]", "", text)
    if not text:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(".", "").replace(",", ".")

    try:
        result = Decimal(text)
    except InvalidOperation:
        return None
    if negative_by_parentheses and result > 0:
        result = -result
    return result


def parse_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _row_number(record: Mapping[str, Any], fallback: int) -> int:
    raw = record.get("__row__", fallback)
    try:
        return int(raw)
    except (TypeError, ValueError):
        return fallback


def _refs(records: Sequence[Mapping[str, Any]]) -> list[_RowRef]:
    return [_RowRef(_row_number(rec, idx), rec) for idx, rec in enumerate(records, start=2)]


def _same_amount(a: Any, b: Any) -> bool:
    left = parse_amount(a)
    right = parse_amount(b)
    if left is None or right is None:
        return left is None and right is None
    return abs(left) == abs(right)


def _same_date(a: Any, b: Any) -> bool:
    left = parse_date(a)
    right = parse_date(b)
    if left is None or right is None:
        return left is None and right is None
    return left == right


def _field_issue(
    code: str,
    label: str,
    left: Any,
    right: Any,
    *,
    compare: str = "text",
) -> ReconciliationIssue | None:
    if compare == "amount":
        equal = _same_amount(left, right)
    elif compare == "date":
        equal = _same_date(left, right)
    elif compare == "compact":
        equal = normalize_compact(left) == normalize_compact(right)
    else:
        equal = normalize_text(left) == normalize_text(right)

    if equal:
        return None
    return ReconciliationIssue(
        code=code,
        message=f"{label} divergente.",
        field=label,
        origin_value=str(left or ""),
        target_value=str(right or ""),
    )


def _matches_exclusion(description: Any, patterns: Iterable[str]) -> bool:
    compact = normalize_compact(description)
    if not compact:
        return False
    for pattern in patterns:
        needle = normalize_compact(pattern)
        if needle and needle in compact:
            return True
    return False


def _process_matches(record: Mapping[str, Any], source_name: str) -> bool:
    process = normalize_compact(record.get("PROCESSO"))
    return not process or process == normalize_compact(source_name)


def _pick_category(item: ReconciliationItem) -> None:
    codes = {issue.code for issue in item.issues}
    if "ORIGEM_SEM_ID" in codes:
        item.category = ReconciliationCategory.ORIGEM_INVALIDA
        item.severity = Severity.CRITICO
        return
    if any(code.startswith("DUP_") for code in codes):
        item.category = ReconciliationCategory.DUPLICADO
        item.severity = Severity.CRITICO
        return
    if "CONTAORDEM_NAO_ENCONTRADA" in codes:
        item.category = ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM
        item.severity = Severity.CRITICO
        return
    if "SOMA_NAO_ENCONTRADO" in codes or "CONTAORDEM_SEM_DOC_SOMA" in codes:
        item.category = ReconciliationCategory.NAO_ENCONTRADO_SOMA
        item.severity = Severity.CRITICO
        return
    if item.issues:
        item.category = ReconciliationCategory.DIVERGENCIA
        item.severity = Severity.MEDIO
        return
    if item.excluded_from_soma:
        item.category = ReconciliationCategory.EXCLUIDO
        item.severity = Severity.INFORMATIVO
        return
    item.category = ReconciliationCategory.OK
    item.severity = Severity.INFORMATIVO


def _doc_rows_with_distinct_ids(rows: Sequence[_RowRef]) -> set[str]:
    return {normalize_doc(r.data.get("ID_INTERNO")) for r in rows if normalize_doc(r.data.get("ID_INTERNO"))}


def reconcile_t_extrato(
    source_records: Sequence[Mapping[str, Any]],
    contaordem_records: Sequence[Mapping[str, Any]],
    soma_records: Sequence[Mapping[str, Any]],
    *,
    year: int,
    month: int,
    exclusions: Sequence[str] = ("ENT.NUMERARIOCH24",),
    source_name: str = "T_EXTRATO",
) -> ReconciliationReport:
    if month < 1 or month > 12:
        raise ValueError("month deve estar entre 1 e 12")

    source_refs_all = _refs(source_records)
    conta_refs = _refs(contaordem_records)
    soma_refs = _refs(soma_records)

    source_refs: list[_RowRef] = []
    outside_period = 0
    invalid_date = 0
    for ref in source_refs_all:
        d = parse_date(ref.data.get("DATA MOV."))
        if d is None:
            invalid_date += 1
            continue
        if d.year != year or d.month != month:
            outside_period += 1
            continue
        source_refs.append(ref)

    source_by_id: dict[str, list[_RowRef]] = defaultdict(list)
    for ref in source_refs:
        source_by_id[normalize_doc(ref.data.get("ID_INTERNO"))].append(ref)

    conta_by_id: dict[str, list[_RowRef]] = defaultdict(list)
    conta_by_doc: dict[str, list[_RowRef]] = defaultdict(list)
    for ref in conta_refs:
        ident = normalize_doc(ref.data.get("ID_INTERNO"))
        if ident:
            conta_by_id[ident].append(ref)
        doc = normalize_doc(ref.data.get("DOC. SOMA"))
        if is_real_doc(doc):
            conta_by_doc[doc].append(ref)

    soma_by_code: dict[str, list[_RowRef]] = defaultdict(list)
    for ref in soma_refs:
        code = normalize_doc(ref.data.get("CODIGO") or ref.data.get("CÓDIGO"))
        if is_real_doc(code):
            soma_by_code[code].append(ref)

    items: list[ReconciliationItem] = []
    for source_ref in source_refs:
        source = source_ref.data
        ident = normalize_doc(source.get("ID_INTERNO"))
        item = ReconciliationItem(source_name=source_name, source_row=source_ref.row, id_interno=ident)
        item.excluded_from_soma = _matches_exclusion(source.get("DESCRIÇÃO"), exclusions)

        if not ident:
            item.issues.append(ReconciliationIssue("ORIGEM_SEM_ID", "Origem sem ID_INTERNO."))
            _pick_category(item)
            items.append(item)
            continue

        if len(source_by_id.get(ident, [])) > 1:
            rows = [r.row for r in source_by_id[ident]]
            item.issues.append(
                ReconciliationIssue(
                    "DUP_ID_ORIGEM",
                    f"ID_INTERNO repetido na origem nas linhas {rows}.",
                    field="ID_INTERNO",
                    origin_value=ident,
                )
            )

        all_candidates = conta_by_id.get(ident, [])
        candidates = [ref for ref in all_candidates if _process_matches(ref.data, source_name)]
        if not candidates and all_candidates:
            candidates = list(all_candidates)

        item.contaordem_rows = [ref.row for ref in candidates]
        if not candidates:
            item.issues.append(
                ReconciliationIssue(
                    "CONTAORDEM_NAO_ENCONTRADA",
                    "ID_INTERNO da origem não foi encontrado na CONTAORDEM.",
                    field="ID_INTERNO",
                    origin_value=ident,
                )
            )
            _pick_category(item)
            items.append(item)
            continue

        if len(candidates) > 1:
            item.issues.append(
                ReconciliationIssue(
                    "DUP_ID_CONTAORDEM",
                    f"ID_INTERNO aparece {len(candidates)} vezes na CONTAORDEM: linhas {item.contaordem_rows}.",
                    field="ID_INTERNO",
                    origin_value=ident,
                )
            )

        conta_ref = candidates[0]
        conta = conta_ref.data

        if not _process_matches(conta, source_name):
            item.issues.append(
                ReconciliationIssue(
                    "PROCESSO_CONTAORDEM_DIVERGENTE",
                    "PROCESSO da CONTAORDEM não corresponde à origem.",
                    field="PROCESSO",
                    origin_value=source_name,
                    target_value=str(conta.get("PROCESSO", "")),
                )
            )

        for issue in (
            _field_issue("DATA_ORIGEM_CONTA", "DATA MOV. origem↔CONTAORDEM", source.get("DATA MOV."), conta.get("DATA MOV."), compare="date"),
            _field_issue("TIPO_ORIGEM_CONTA", "TIPO origem↔CONTAORDEM", source.get("TIPO"), conta.get("TIPO"), compare="compact"),
            _field_issue("DESCRICAO_ORIGEM_CONTA", "DESCRIÇÃO origem↔CONTAORDEM", source.get("DESCRIÇÃO"), conta.get("DESCRIÇÃO"), compare="compact"),
            _field_issue("VALOR_ORIGEM_CONTA", "VALOR origem↔CONTAORDEM", source.get("IMPORTÂNCIA"), conta.get("IMPORTÂNCIA"), compare="amount"),
        ):
            if issue:
                item.issues.append(issue)

        source_doc = normalize_doc(source.get("DOC. SOMA"))
        conta_doc = normalize_doc(conta.get("DOC. SOMA"))
        item.doc_soma = conta_doc if is_real_doc(conta_doc) else source_doc if is_real_doc(source_doc) else conta_doc

        if is_real_doc(source_doc) and is_real_doc(conta_doc) and source_doc != conta_doc:
            item.issues.append(
                ReconciliationIssue(
                    "DOC_ORIGEM_CONTA",
                    "DOC. SOMA da origem diverge da CONTAORDEM.",
                    field="DOC. SOMA",
                    origin_value=source_doc,
                    target_value=conta_doc,
                )
            )
        elif not is_real_doc(source_doc) and is_real_doc(conta_doc):
            placeholder = normalize_compact(source_doc)
            if placeholder in _PLACEHOLDER_DOCS or not placeholder:
                item.warnings.append(
                    ReconciliationIssue(
                        "ORIGEM_DOC_PENDENTE",
                        "A origem ainda não contém o DOC. SOMA definitivo, mas a CONTAORDEM já contém.",
                        field="DOC. SOMA",
                        origin_value=source_doc,
                        target_value=conta_doc,
                    )
                )

        if is_real_doc(conta_doc):
            doc_rows = conta_by_doc.get(conta_doc, [])
            distinct_ids = _doc_rows_with_distinct_ids(doc_rows)
            if len(distinct_ids) > 1:
                item.issues.append(
                    ReconciliationIssue(
                        "DUP_DOC_CONTAORDEM",
                        f"DOC. SOMA {conta_doc} está associado a vários IDs na CONTAORDEM: {sorted(distinct_ids)}.",
                        field="DOC. SOMA",
                        origin_value=conta_doc,
                    )
                )

            if item.excluded_from_soma:
                _pick_category(item)
                items.append(item)
                continue
        elif item.excluded_from_soma:
            _pick_category(item)
            items.append(item)
            continue
        else:
            item.issues.append(
                ReconciliationIssue(
                    "CONTAORDEM_SEM_DOC_SOMA",
                    "CONTAORDEM não possui um DOC. SOMA numérico para localizar o lançamento no SOMA.",
                    field="DOC. SOMA",
                    target_value=str(conta.get("DOC. SOMA", "")),
                )
            )
            _pick_category(item)
            items.append(item)
            continue

        soma_candidates = soma_by_code.get(conta_doc, [])
        item.soma_rows = [ref.row for ref in soma_candidates]
        if not soma_candidates:
            item.issues.append(
                ReconciliationIssue(
                    "SOMA_NAO_ENCONTRADO",
                    f"DOC. SOMA {conta_doc} não foi encontrado na sheet SOMA.",
                    field="CODIGO",
                    origin_value=conta_doc,
                )
            )
            _pick_category(item)
            items.append(item)
            continue

        if len(soma_candidates) > 1:
            item.issues.append(
                ReconciliationIssue(
                    "DUP_CODIGO_SOMA",
                    f"CODIGO {conta_doc} aparece {len(soma_candidates)} vezes na sheet SOMA: linhas {item.soma_rows}.",
                    field="CODIGO",
                    origin_value=conta_doc,
                )
            )

        soma = soma_candidates[0].data
        for issue in (
            _field_issue("TIPO_CONTA_SOMA", "TIPO CONTAORDEM↔SOMA", conta.get("TIPO"), soma.get("TIPO"), compare="compact"),
            _field_issue("DESCRICAO_CONTA_SOMA", "DESCRIÇÃO CONTAORDEM↔SOMA", conta.get("DESCRIÇÃO SOMA"), soma.get("DESCRIÇÃO"), compare="compact"),
            _field_issue("VALOR_CONTA_SOMA", "VALOR CONTAORDEM↔SOMA", conta.get("IMPORTÂNCIA"), soma.get("VALOR"), compare="amount"),
            _field_issue("DATA_CONTA_SOMA", "DATA CONTAORDEM↔SOMA", conta.get("DATA MOV."), soma.get("PAGAMENTO"), compare="date"),
        ):
            if issue:
                item.issues.append(issue)

        status_soma = normalize_compact(soma.get("STATUS"))
        if status_soma and status_soma != "PAGO":
            item.warnings.append(
                ReconciliationIssue(
                    "SOMA_STATUS_NAO_PAGO",
                    "Lançamento existe no SOMA, mas STATUS não é PAGO.",
                    field="STATUS",
                    target_value=str(soma.get("STATUS", "")),
                )
            )
        baixa = normalize_compact(soma.get("BAIXA"))
        if baixa and baixa != "SIM":
            item.warnings.append(
                ReconciliationIssue(
                    "SOMA_BAIXA_NAO_SIM",
                    "Lançamento existe no SOMA, mas BAIXA não é SIM.",
                    field="BAIXA",
                    target_value=str(soma.get("BAIXA", "")),
                )
            )

        _pick_category(item)
        items.append(item)

    return ReconciliationReport(
        source_name=source_name,
        year=year,
        month=month,
        items=items,
        source_rows_total=len(source_refs),
        source_rows_outside_period=outside_period,
        source_rows_invalid_date=invalid_date,
    )
