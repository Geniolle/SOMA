from __future__ import annotations

import argparse
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence
from zoneinfo import ZoneInfo

from dotenv import load_dotenv

from soma_app.domain.reconciliation import (
    ReconciliationCategory,
    ReconciliationItem,
    ReconciliationReport,
    reconcile_t_extrato,
)
from soma_app.infra.env import env_str
from soma_app.infra.sheets_client import SheetsClient

DEFAULT_EXCLUSIONS = ("ENT.NUMERARIOCH24",)


# Nomes de mês aceites em português (case-insensitive, com e sem acento).
_MONTH_PT: dict[str, int] = {
    "JANEIRO": 1, "JAN": 1,
    "FEVEREIRO": 2, "FEV": 2,
    "MARCO": 3, "MARÇO": 3, "MAR": 3,
    "ABRIL": 4, "ABR": 4,
    "MAIO": 5, "MAI": 5,
    "JUNHO": 6, "JUN": 6,
    "JULHO": 7, "JUL": 7,
    "AGOSTO": 8, "AGO": 8,
    "SETEMBRO": 9, "SET": 9,
    "OUTUBRO": 10, "OUT": 10,
    "NOVEMBRO": 11, "NOV": 11,
    "DEZEMBRO": 12, "DEZ": 12,
}


def _parse_month(raw: str) -> int:
    """Converte o valor de B1 em número de mês (1-12).

    Aceita:
    - Número inteiro como string: ``"4"``
    - Nome completo em português: ``"ABRIL"``, ``"abril"``
    - Abreviação de 3 letras: ``"ABR"``
    - Com ou sem acento: ``"MARÇO"`` == ``"MARCO"``
    """
    text = raw.strip().upper()
    text_norm = "".join(
        ch for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )
    if text_norm.isdigit():
        m = int(text_norm)
        if 1 <= m <= 12:
            return m
        raise ValueError(f"Mês numérico fora do intervalo 1-12: {raw!r}")
    # tentar nome original e normalizado
    for candidate in (text, text_norm):
        if candidate in _MONTH_PT:
            return _MONTH_PT[candidate]
        # abreviatura de 3 letras
        if candidate[:3] in _MONTH_PT:
            return _MONTH_PT[candidate[:3]]
    raise ValueError(
        f"Não foi possível interpretar o mês a partir de B1: {raw!r}. "
        f"Use um número (1-12) ou o nome em português (ex.: ABRIL)."
    )


def _read_period_from_sheet(sheets: SheetsClient, ws: str) -> tuple[int, int]:
    """Lê o período (ano, mês) das células A1 e B1 da sheet ``ws``.

    Formato esperado:
    - **A1** — Ano (inteiro, ex.: ``2026``)
    - **B1** — Mês (número ``4`` ou nome em português ``ABRIL``)
    """
    row1 = sheets._ws(ws).row_values(1)  # devolve lista de strings
    if len(row1) < 2 or not row1[0].strip() or not row1[1].strip():
        raise ValueError(
            f"Sheet {ws!r}: A1 e/ou B1 estão vazios. "
            "A1 deve conter o ano e B1 o mês (número ou nome em português)."
        )
    try:
        year = int(row1[0].strip())
    except ValueError:
        raise ValueError(
            f"Sheet {ws!r}: A1={row1[0]!r} não é um ano válido."
        ) from None
    month = _parse_month(row1[1])
    return year, month


def _resolve_period_cli(year: int | None, month: int | None) -> tuple[int, int] | None:
    """Valida e devolve o período fornecido pela CLI, ou ``None`` se omitido."""
    if year is None and month is None:
        return None
    if (year is None) != (month is None):
        raise ValueError("Informe --year e --month juntos, ou omita ambos.")
    if month is not None and (month < 1 or month > 12):
        raise ValueError("--month deve estar entre 1 e 12.")
    return year, month  # type: ignore[return-value]


# Raiz do projecto calculada a partir da localização deste ficheiro.
# src/soma_app/workflows/reconcile_t_extrato.py → 3 níveis acima → raiz
_PROJECT_ROOT = Path(__file__).resolve().parents[3]


def _load_env(env_file: str | None) -> Path:
    """Carrega o ficheiro .env com a seguinte ordem de prioridade:

    1. ``--env-file`` passado na CLI (ou env var ``ENV_FILE``).
    2. ``deploy/.env`` relativo ao directório de trabalho actual (CWD).
    3. ``deploy/.env`` relativo à raiz do projecto (calculada a partir do
       caminho deste ficheiro — funciona mesmo que o CWD seja diferente).

    Desta forma, correr o script directamente (``python reconcile_t_extrato.py``)
    sem argumentos encontra as credenciais automaticamente.
    """
    explicit = env_file or env_str("ENV_FILE", "")

    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    # fallbacks — procura deploy/.env em dois pontos de referência
    candidates += [
        Path("deploy") / ".env",            # relativo ao CWD
        _PROJECT_ROOT / "deploy" / ".env",  # relativo à raiz do projecto
    ]

    for path in candidates:
        if path.exists():
            load_dotenv(dotenv_path=path, override=False)
            return path.resolve()

    # Nenhum ficheiro encontrado — usa o primeiro candidato (pode ser explícito)
    # e deixa load_dotenv falhar silenciosamente; o erro real surgirá depois
    # quando as credenciais forem validadas.
    fallback = candidates[0]
    load_dotenv(dotenv_path=fallback, override=False)
    return fallback.resolve()


def _sheet_client_settings() -> dict[str, str]:
    creds = env_str("GOOGLE_CREDENTIALS_PATH") or env_str("GOOGLE_APPLICATION_CREDENTIALS")
    spreadsheet_url = env_str("SPREADSHEET_URL")
    spreadsheet_id = env_str("SPREADSHEET_ID")
    spreadsheet_name = env_str("SPREADSHEET_NAME") or env_str("SPREADSHEET")

    if not creds:
        raise RuntimeError("GOOGLE_CREDENTIALS_PATH/GOOGLE_APPLICATION_CREDENTIALS não configurado.")
    if not (spreadsheet_url or spreadsheet_id or spreadsheet_name):
        raise RuntimeError("Configure SPREADSHEET_URL, SPREADSHEET_ID ou SPREADSHEET_NAME.")

    return {
        "google_credentials_path": creds,
        "spreadsheet_url": spreadsheet_url,
        "spreadsheet_id": spreadsheet_id,
        "spreadsheet_name": spreadsheet_name,
    }


def _with_rows(records: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for row, rec in enumerate(records, start=2):
        item = dict(rec)
        item["__row__"] = row
        out.append(item)
    return out


def _read_records(sheets: SheetsClient, ws: str) -> list[dict[str, Any]]:
    """Lê registos preservando valores formatados (ex.: ``230,00`` em vez de
    ``23000``). Usa :meth:`~SheetsClient.get_all_records_formatted` para
    impedir que o gspread converta automaticamente strings numéricas.
    """
    return _with_rows(sheets.get_all_records_formatted(ws))


def _parse_exclusions(cli_values: Sequence[str] | None) -> tuple[str, ...]:
    if cli_values:
        return tuple(v for v in cli_values if str(v).strip())
    env_value = env_str("RECONCILE_T_EXTRATO_EXCLUSIONS")
    if env_value:
        values = [part.strip() for part in env_value.replace(";", ",").split(",")]
        return tuple(value for value in values if value)
    return DEFAULT_EXCLUSIONS


def _print_load_header(source: str, year: int, month: int) -> None:
    print()
    print("=" * 78)
    print(f"CONCILIAÇÃO READ-ONLY | {source} → CONTAORDEM → SOMA | {month:02d}/{year}")
    print("=" * 78)
    print("Nenhuma célula será alterada. O processo apenas lê as sheets e gera o reporte.")
    print()


def _count_downstream(report: ReconciliationReport) -> tuple[int, int]:
    conta = sum(bool(item.contaordem_rows) for item in report.items)
    soma = sum(bool(item.soma_rows) for item in report.items)
    return conta, soma


def _print_summary(report: ReconciliationReport) -> None:
    counts = report.category_counts()
    conta_found, soma_found = _count_downstream(report)
    print()
    print("-" * 78)
    print("RESUMO")
    print("-" * 78)
    print(f"Origem analisada........................ {report.source_rows_total}")
    print(f"Encontrados na CONTAORDEM............... {conta_found}")
    print(f"Encontrados no SOMA..................... {soma_found}")
    print(f"✅ Trilogia OK.......................... {counts[ReconciliationCategory.OK]}")
    print(f"⏭  Excluídos intencionalmente do SOMA... {counts[ReconciliationCategory.EXCLUIDO]}")
    print(f"❌ Não encontrados na CONTAORDEM........ {counts[ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM]}")
    print(f"❌ Não encontrados no SOMA.............. {counts[ReconciliationCategory.NAO_ENCONTRADO_SOMA]}")
    print(f"⚠  Divergências......................... {counts[ReconciliationCategory.DIVERGENCIA]}")
    print(f"🚨 Duplicados............................ {counts[ReconciliationCategory.DUPLICADO]}")
    print(f"🚨 Origem inválida....................... {counts[ReconciliationCategory.ORIGEM_INVALIDA]}")
    print(f"🟡 Avisos adicionais..................... {report.warning_count}")
    print(f"Taxa de conciliação...................... {report.reconciliation_rate:.2f}%")
    if report.source_rows_invalid_date:
        print(f"Linhas ignoradas por DATA MOV. inválida.. {report.source_rows_invalid_date}")
    print("-" * 78)


def _item_rows(item: ReconciliationItem) -> str:
    parts = [f"origem={item.source_row}"]
    if item.contaordem_rows:
        parts.append("CONTAORDEM=" + ",".join(str(row) for row in item.contaordem_rows))
    if item.soma_rows:
        parts.append("SOMA=" + ",".join(str(row) for row in item.soma_rows))
    return " | ".join(parts)


def _print_item(item: ReconciliationItem) -> None:
    ident = item.id_interno or "<SEM ID>"
    doc = item.doc_soma or "-"
    print(f"[{item.severity.value}] ID={ident} | DOC.SOMA={doc} | {_item_rows(item)}")
    print(f"  FUNIL: {item.flow}")
    for issue in item.issues:
        detail = issue.message
        if issue.field and (issue.origin_value or issue.target_value):
            detail += f" [{issue.field}: {issue.origin_value or '-'} ≠ {issue.target_value or '-'}]"
        print(f"  - {detail}")
    for warning in item.warnings:
        detail = warning.message
        if warning.field and (warning.origin_value or warning.target_value):
            detail += f" [{warning.field}: {warning.origin_value or '-'} → {warning.target_value or '-'}]"
        print(f"  ! AVISO: {detail}")
    print(f"  AÇÃO: {item.action}")
    print()


def _print_section(title: str, items: Sequence[ReconciliationItem]) -> None:
    if not items:
        return
    print()
    print("=" * 78)
    print(f"{title} ({len(items)})")
    print("=" * 78)
    for item in items:
        _print_item(item)


def print_report(
    report: ReconciliationReport,
    *,
    show_ok: bool = False,
    show_excluded: bool = False,
    show_warnings: bool = True,
) -> None:
    _print_summary(report)

    by_category: dict[ReconciliationCategory, list[ReconciliationItem]] = {
        category: [] for category in ReconciliationCategory
    }
    for item in report.items:
        by_category[item.category].append(item)

    _print_section("🚨 DUPLICADOS", by_category[ReconciliationCategory.DUPLICADO])
    _print_section(
        "❌ NÃO ENCONTRADOS NA CONTAORDEM",
        by_category[ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM],
    )
    _print_section("❌ NÃO ENCONTRADOS NO SOMA", by_category[ReconciliationCategory.NAO_ENCONTRADO_SOMA])
    _print_section("⚠ DIVERGÊNCIAS", by_category[ReconciliationCategory.DIVERGENCIA])
    _print_section("🚨 ORIGEM INVÁLIDA", by_category[ReconciliationCategory.ORIGEM_INVALIDA])

    if show_warnings:
        warning_items = [item for item in report.items if item.warnings and item.category == ReconciliationCategory.OK]
        _print_section("🟡 AVISOS EM REGISTOS CONCILIADOS", warning_items)

    if show_excluded:
        _print_section("⏭ EXCLUÍDOS DO SOMA", by_category[ReconciliationCategory.EXCLUIDO])

    if show_ok:
        ok_without_warning = [item for item in by_category[ReconciliationCategory.OK] if not item.warnings]
        _print_section("✅ TRILOGIA OK", ok_without_warning)

    problems = sum(
        len(by_category[category])
        for category in (
            ReconciliationCategory.DUPLICADO,
            ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM,
            ReconciliationCategory.NAO_ENCONTRADO_SOMA,
            ReconciliationCategory.DIVERGENCIA,
            ReconciliationCategory.ORIGEM_INVALIDA,
        )
    )
    print()
    if problems:
        print(f"RESULTADO: {problems} registo(s) precisam de tratamento.")
    else:
        print("RESULTADO: nenhum registo crítico/divergente encontrado no período.")
    print()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Conciliação read-only T_EXTRATO → CONTAORDEM → SOMA com reporte no terminal."
    )
    parser.add_argument("--year", type=int, help="Ano a analisar. Opcional: se omitido, lê A1 da sheet de período.")
    parser.add_argument("--month", type=int, help="Mês (1-12). Opcional: se omitido, lê B1 da sheet de período (aceita número ou nome, ex.: ABRIL).")
    parser.add_argument("--env-file", help="Caminho do .env. Default: ENV_FILE ou deploy/.env.")
    parser.add_argument("--source-sheet", default="T_EXTRATO", help="Nome da sheet de origem. Default: T_EXTRATO.")
    parser.add_argument("--period-sheet", help="Sheet com A1=ano e B1=mês. Default: SHEET_PERIOD ou T_AUDIT.")
    parser.add_argument("--contaordem-sheet", help="Nome da CONTAORDEM. Default: SHEET_CONTAORDEM/CONTAORDEM.")
    parser.add_argument("--soma-sheet", help="Nome da SOMA. Default: SHEET_SOMA/SOMA.")
    parser.add_argument(
        "--exclude-description",
        action="append",
        help="Trecho de descrição a excluir da obrigação de existir no SOMA. Pode repetir.",
    )
    parser.add_argument("--show-ok", action="store_true", help="Mostra também todos os registos com trilogia OK.")
    parser.add_argument("--show-excluded", action="store_true", help="Mostra os registos excluídos intencionalmente do SOMA.")
    parser.add_argument("--no-warnings", action="store_true", help="Não imprime a secção de avisos em registos conciliados.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Executa a conciliação sem gravar os resultados na coluna STATUS da Google Sheet.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        cli_period = _resolve_period_cli(args.year, args.month)
        env_path = _load_env(args.env_file)
        sheets = SheetsClient(_sheet_client_settings())
    except Exception as exc:
        print(f"ERRO DE CONFIGURAÇÃO: {exc}")
        return 2

    source_sheet = str(args.source_sheet or "T_EXTRATO").strip()
    if source_sheet.upper() != "T_EXTRATO":
        print("ERRO: esta primeira versão suporta apenas a origem T_EXTRATO.")
        return 2

    contaordem_sheet = str(args.contaordem_sheet or env_str("SHEET_CONTAORDEM", "CONTAORDEM")).strip()
    soma_sheet = str(args.soma_sheet or env_str("SHEET_SOMA", "SOMA")).strip()
    period_sheet = str(getattr(args, "period_sheet", None) or env_str("SHEET_PERIOD", "T_AUDIT")).strip()
    exclusions = _parse_exclusions(args.exclude_description)

    # Resolver período: CLI > sheet A1/B1 (T_AUDIT) > data atual
    if cli_period is not None:
        year, month = cli_period
        period_source = "argumento --year/--month"
    else:
        try:
            year, month = _read_period_from_sheet(sheets, period_sheet)
            period_source = f"células A1/B1 da sheet {period_sheet!r}"
        except Exception as exc:
            now = datetime.now(ZoneInfo("Europe/Lisbon"))
            year, month = now.year, now.month
            period_source = f"data atual (não foi possível ler A1/B1: {exc})"

    try:
        _print_load_header(source_sheet, year, month)
        print(f"Configuração carregada de: {env_path}")
        print(f"Período lido de: {period_source}")
        print(f"Exclusões do SOMA: {', '.join(exclusions) if exclusions else '(nenhuma)'}")
        print()
        print(f"(1/3) Lendo origem {source_sheet}...")
        source = _read_records(sheets, source_sheet)
        print(f"      {len(source)} linhas carregadas.")

        print(f"(2/3) Lendo {contaordem_sheet}...")
        contaordem = _read_records(sheets, contaordem_sheet)
        print(f"      {len(contaordem)} linhas carregadas.")

        print(f"(3/3) Lendo {soma_sheet}...")
        soma = _read_records(sheets, soma_sheet)
        print(f"      {len(soma)} linhas carregadas.")
    except Exception as exc:
        print(f"ERRO DE LEITURA: {type(exc).__name__}: {exc}")
        return 3

    report = reconcile_t_extrato(
        source,
        contaordem,
        soma,
        year=year,
        month=month,
        exclusions=exclusions,
        source_name=source_sheet,
    )
    print_report(
        report,
        show_ok=args.show_ok,
        show_excluded=args.show_excluded,
        show_warnings=not args.no_warnings,
    )

    # Escrever os resultados na coluna STATUS da sheet T_EXTRATO
    if not args.dry_run:
        updates = [
            {"row": item.source_row, "STATUS": item.status_label}
            for item in report.items
        ]
        if updates:
            print("-" * 78)
            print(f"Gravando coluna STATUS na sheet '{source_sheet}' ({len(updates)} linhas)...")
            try:
                sheets.batch_update_rows(source_sheet, updates)
                print("✅ Gravado com sucesso na Google Sheet!")
            except Exception as exc:
                print(f"❌ ERRO AO GRAVAR NA GOOGLE SHEET: {exc}")
            print("-" * 78)
    else:
        print("-" * 78)
        print("⏭ MODO --dry-run ATIVO: Nenhuma alteração foi gravada na Google Sheet.")
        print("-" * 78)

    critical = Counter(item.category for item in report.items)
    has_problems = any(
        critical[category] > 0
        for category in (
            ReconciliationCategory.DUPLICADO,
            ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM,
            ReconciliationCategory.NAO_ENCONTRADO_SOMA,
            ReconciliationCategory.DIVERGENCIA,
            ReconciliationCategory.ORIGEM_INVALIDA,
        )
    )
    return 1 if has_problems else 0


if __name__ == "__main__":
    raise SystemExit(main())

