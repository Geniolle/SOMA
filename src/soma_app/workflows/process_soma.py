from __future__ import annotations

import logging
import re
import time
import unicodedata
from calendar import monthrange
from datetime import datetime
from typing import Any, Dict, List, Tuple

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.infra.trace import step, log_kv
from soma_app.workflows.process_contaordem import SheetsTable

log = logging.getLogger("soma_app.workflows.process_soma")

TABLE_RESULTADOS = (By.CSS_SELECTOR, "table#exampleTableSearch")
HEADER_CELLS = (By.CSS_SELECTOR, "table#exampleTableSearch thead th")
BODY_ROWS = (By.CSS_SELECTOR, "table#exampleTableSearch tbody tr")
EMPTY_ROW = (By.CSS_SELECTOR, "table#exampleTableSearch tbody td.dataTables_empty")


def _get_setting(settings: Any, *names: str, default: str = "") -> str:
    for name in names:
        try:
            if isinstance(settings, dict):
                v = settings.get(name)
            else:
                v = getattr(settings, name, None)
        except Exception:
            v = None
        if v is None:
            continue
        s = str(v).strip()
        if s:
            return s
    return default


def _norm_text(v: Any) -> str:
    s = "" if v is None else str(v).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.lower().split())


def _norm_code(v: Any) -> str:
    s = "" if v is None else str(v).strip()
    if not s:
        return ""
    m = re.fullmatch(r"(\d+)\.0+", s)
    if m:
        return m.group(1)
    return s


def _now_pt() -> datetime:
    try:
        from zoneinfo import ZoneInfo

        return datetime.now(ZoneInfo("Europe/Lisbon"))
    except Exception:
        return datetime.now()


def _intervalo_mes_anterior_ate_mes_atual() -> Tuple[str, str]:
    now = _now_pt()

    if now.month == 1:
        prev_month = 12
        prev_year = now.year - 1
    else:
        prev_month = now.month - 1
        prev_year = now.year

    start_prev_month = datetime(prev_year, prev_month, 1)
    end_current_month = datetime(now.year, now.month, monthrange(now.year, now.month)[1])

    return start_prev_month.strftime("%d/%m/%Y"), end_current_month.strftime("%d/%m/%Y")


def _sheet_soma_name(settings: Any) -> str:
    return _get_setting(settings, "sheet_soma", "SHEET_SOMA", default="SOMA")


def _resolve_col_name(table: SheetsTable, *aliases: str) -> str:
    by_norm = {_norm_text(h): h for h in table.header}
    for alias in aliases:
        hit = by_norm.get(_norm_text(alias))
        if hit:
            return hit
    raise RuntimeError(
        f"Coluna não encontrada na sheet. aliases={aliases} | header={table.header}"
    )


def _set_input_value(actions: Any, locator: Tuple[str, str], value: str, timeout_seconds: int = 20) -> None:
    el = actions.wait_present(locator, timeout_seconds=timeout_seconds)

    try:
        actions.scroll_into_view(locator)
    except Exception:
        pass

    try:
        actions.click_js(locator)
    except Exception:
        pass

    try:
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.DELETE)
    except Exception:
        try:
            el.clear()
        except Exception:
            pass

    try:
        el.send_keys(value)
    except Exception:
        actions.driver.execute_script(
            """
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """,
            el,
            value,
        )

    try:
        actions.driver.execute_script(
            """
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """,
            el,
        )
    except Exception:
        pass


def _abrir_listagem(page: EntradasSaidasPage) -> None:
    if page.a.exists(page.PESQ_DESCRICAO, timeout_seconds=2):
        return

    page._dismiss_overlays()
    page._close_datepicker()
    page._click_menu_entradas_saidas(timeout_seconds=60)
    page.a.wait_present(page.PESQ_DESCRICAO, timeout_seconds=30)


def _mapear_colunas_resultado(actions: Any) -> Dict[str, int]:
    headers = actions.driver.find_elements(*HEADER_CELLS)
    mapa: Dict[str, int] = {}

    for idx, th in enumerate(headers):
        nome = _norm_text(th.text)
        if nome:
            mapa[nome] = idx

    return mapa


def _esperar_resultado_pesquisa(actions: Any, timeout_seconds: int = 30) -> None:
    t0 = time.time()
    while time.time() - t0 <= timeout_seconds:
        try:
            if actions.exists(EMPTY_ROW, timeout_seconds=1):
                return
        except Exception:
            pass

        try:
            rows = actions.driver.find_elements(*BODY_ROWS)
            for tr in rows:
                tds = tr.find_elements(By.TAG_NAME, "td")
                if tds:
                    return
        except Exception:
            pass

        time.sleep(0.4)

    raise TimeoutException("Timeout à espera do resultado da pesquisa de lançamentos do SOMA.")


def _ler_lancamentos_tabela(actions: Any) -> List[Dict[str, str]]:
    actions.wait_present(TABLE_RESULTADOS, timeout_seconds=30)

    if actions.exists(EMPTY_ROW, timeout_seconds=2):
        return []

    mapa = _mapear_colunas_resultado(actions)
    obrigatorias = {
        "codigo": "codigo",
        "tipo": "tipo",
        "descricao": "descricao",
        "valor": "valor",
        "pagamento": "pagamento",
        "status": "status",
        "baixa": "baixa",
    }

    faltantes = [chave for chave, alias in obrigatorias.items() if alias not in mapa]
    if faltantes:
        raise RuntimeError(
            f"Colunas obrigatórias não encontradas na tabela do SOMA: {faltantes} | mapa={mapa}"
        )

    rows = actions.driver.find_elements(*BODY_ROWS)
    resultados: List[Dict[str, str]] = []

    for tr in rows:
        tds = tr.find_elements(By.TAG_NAME, "td")
        if not tds:
            continue

        max_idx = max(mapa[a] for a in obrigatorias.values())
        if len(tds) <= max_idx:
            continue

        codigo = _norm_code(tds[mapa["codigo"]].text)
        if not codigo:
            continue

        item = {
            "codigo": codigo,
            "tipo": (tds[mapa["tipo"]].text or "").strip(),
            "descricao": (tds[mapa["descricao"]].text or "").strip(),
            "valor": (tds[mapa["valor"]].text or "").strip(),
            "pagamento": (tds[mapa["pagamento"]].text or "").strip(),
            "status": (tds[mapa["status"]].text or "").strip(),
            "baixa": (tds[mapa["baixa"]].text or "").strip(),
        }
        resultados.append(item)

    return resultados


def _pesquisar_lancamentos_periodo(page: EntradasSaidasPage, data_ini: str, data_fim: str) -> List[Dict[str, str]]:
    with step(log, "soma.search_site", data_ini=data_ini, data_fim=data_fim):
        _abrir_listagem(page)
        page._dismiss_overlays()
        page._close_datepicker()

        _set_input_value(page.a, page.PESQ_DESCRICAO, "")
        page.a.click_js(page.RADIO_PERIODO)
        time.sleep(0.3)
        page.a.click_js(page.RADIO_DATA_PAGAMENTO)
        time.sleep(0.3)

        _set_input_value(page.a, page.DATA_INI, data_ini)
        _set_input_value(page.a, page.DATA_FIM, data_fim)
        page._close_datepicker()

        print(f"(7.1) Pesquisando lançamentos no período: {data_ini} até {data_fim}")
        page.a.click_js(page.BTN_PESQUISAR)
        page._dismiss_overlays()

        _esperar_resultado_pesquisa(page.a, timeout_seconds=30)
        resultados = _ler_lancamentos_tabela(page.a)

        print(f"(7.2) Total de lançamentos encontrados no site: {len(resultados)}")
        log_kv(log, logging.INFO, "Lançamentos carregados do SOMA.", data_ini=data_ini, data_fim=data_fim, total=len(resultados))
        return resultados


def _existing_codes(table: SheetsTable, code_col: str) -> set[str]:
    out: set[str] = set()
    for rec in table.records:
        codigo = _norm_code(rec.get(code_col))
        if codigo:
            out.add(codigo)
    return out


def _is_target_row_blank(rec: Dict[str, Any], target_cols: List[str]) -> bool:
    return all(not str(rec.get(col, "") or "").strip() for col in target_cols)


def _next_write_row(table: SheetsTable, code_col: str, target_cols: List[str]) -> int:
    first_blank_row: int | None = None

    for row_idx, rec in enumerate(table.records, start=2):
        if _is_target_row_blank(rec, target_cols):
            first_blank_row = row_idx
            break

    if first_blank_row is None:
        return len(table.records) + 2

    for row_idx, rec in enumerate(table.records, start=2):
        if row_idx <= first_blank_row:
            continue
        if _norm_code(rec.get(code_col)):
            log_kv(
                log,
                logging.WARNING,
                "Existe buraco interno na sheet SOMA. Para evitar sobregravar dados, vou escrever no final.",
                first_blank_row=first_blank_row,
                row_with_data_below=row_idx,
            )
            return len(table.records) + 2

    return first_blank_row


def _filtrar_novos_lancamentos(lancamentos: List[Dict[str, str]], existing_codes: set[str]) -> Tuple[List[Dict[str, str]], int, int]:
    novos: List[Dict[str, str]] = []
    ignorados_existente = 0
    ignorados_lote = 0
    seen_batch: set[str] = set()

    for item in lancamentos:
        codigo = _norm_code(item.get("codigo"))
        if not codigo:
            continue

        if codigo in existing_codes:
            ignorados_existente += 1
            continue

        if codigo in seen_batch:
            ignorados_lote += 1
            continue

        seen_batch.add(codigo)
        novos.append(
            {
                "codigo": codigo,
                "tipo": item.get("tipo", ""),
                "descricao": item.get("descricao", ""),
                "valor": item.get("valor", ""),
                "pagamento": item.get("pagamento", ""),
                "status": item.get("status", ""),
                "baixa": item.get("baixa", ""),
            }
        )

    return novos, ignorados_existente, ignorados_lote


def atualizar_sheet_soma(sheets: Any, actions: Any, settings: Any) -> None:
    ws_soma = _sheet_soma_name(settings)
    data_ini, data_fim = _intervalo_mes_anterior_ate_mes_atual()

    print()
    print("==================================================================")
    print("(7) Iniciando o processo 'Atualizar sheet SOMA'")
    print("==================================================================")
    print()

    page = EntradasSaidasPage(actions, settings)

    with step(log, "soma.run", ws=ws_soma, data_ini=data_ini, data_fim=data_fim):
        lancamentos = _pesquisar_lancamentos_periodo(page, data_ini, data_fim)

        table = SheetsTable(sheets, ws_soma)
        table.load()

        col_codigo = _resolve_col_name(table, "CODIGO", "CÓDIGO")
        col_tipo = _resolve_col_name(table, "TIPO")
        col_descricao = _resolve_col_name(table, "DESCRICAO", "DESCRIÇÃO")
        col_valor = _resolve_col_name(table, "VALOR")
        col_pagamento = _resolve_col_name(table, "PAGAMENTO")
        col_status = _resolve_col_name(table, "STATUS")
        col_baixa = _resolve_col_name(table, "BAIXA")

        target_cols = [
            col_codigo,
            col_tipo,
            col_descricao,
            col_valor,
            col_pagamento,
            col_status,
            col_baixa,
        ]

        codigos_existentes = _existing_codes(table, col_codigo)
        print(f"(7.3) Códigos já existentes na sheet SOMA: {len(codigos_existentes)}")

        novos, ignorados_existente, ignorados_lote = _filtrar_novos_lancamentos(lancamentos, codigos_existentes)
        print(f"(7.4) Lançamentos novos para inserir: {len(novos)}")
        print(f"(7.5) Ignorados por já existirem na sheet: {ignorados_existente}")
        print(f"(7.6) Ignorados por repetição dentro do lote: {ignorados_lote}")

        if not novos:
            log_kv(
                log,
                logging.INFO,
                "Nenhum lançamento novo para inserir na sheet SOMA.",
                ws=ws_soma,
                encontrados=len(lancamentos),
                existentes=len(codigos_existentes),
                ignorados_existente=ignorados_existente,
                ignorados_lote=ignorados_lote,
            )
            print("(7.7) Nenhum lançamento novo encontrado. Sheet SOMA já está atualizada.\n")
            return

        start_row = _next_write_row(table, col_codigo, target_cols)
        print(f"(7.7) Primeira linha segura para escrita na sheet SOMA: {start_row}")

        updates: List[Tuple[int, str, Any]] = []
        inserted_codes: List[str] = []

        for offset, item in enumerate(novos):
            row_idx = start_row + offset
            inserted_codes.append(item["codigo"])
            updates.extend(
                [
                    (row_idx, col_codigo, item["codigo"]),
                    (row_idx, col_tipo, item["tipo"]),
                    (row_idx, col_descricao, item["descricao"]),
                    (row_idx, col_valor, item["valor"]),
                    (row_idx, col_pagamento, item["pagamento"]),
                    (row_idx, col_status, item["status"]),
                    (row_idx, col_baixa, item["baixa"]),
                ]
            )

        table.batch_update_cells(updates)

        log_kv(
            log,
            logging.INFO,
            "Sheet SOMA atualizada com sucesso.",
            ws=ws_soma,
            data_ini=data_ini,
            data_fim=data_fim,
            encontrados=len(lancamentos),
            existentes=len(codigos_existentes),
            inseridos=len(novos),
            ignorados_existente=ignorados_existente,
            ignorados_lote=ignorados_lote,
            start_row=start_row,
            codigos=", ".join(inserted_codes[:100]),
        )

        print(f"(7.8) Sheet '{ws_soma}' atualizada com sucesso. Inseridos: {len(novos)}")
        print("(7.9) Códigos inseridos:")
        for codigo in inserted_codes:
            print(f" - {codigo}")
        print()