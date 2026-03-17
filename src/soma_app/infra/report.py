# src/soma_app/infra/report.py
from __future__ import annotations

import logging
import os
from typing import Any, Dict, Optional, Tuple


# Estilo do report:
# - legacy (default): imprime blocos e mensagens humanas
# - off: desliga o report
_STYLE = (os.getenv("SOMA_REPORT_STYLE") or "legacy").strip().lower()
_ENABLED = _STYLE not in {"0", "false", "off", "none"}

_REPORT_LOGGER_NAME = "soma_report"
_logger = logging.getLogger(_REPORT_LOGGER_NAME)

_state: Dict[str, Any] = {
    "printed": set(),     # ids de secções/subsecções já impressas
    "sheet": None,        # nome da sheet (se aparecer nos fields)
    "current_row": None,  # linha atual (se aparecer)
    "current_tipo": None, # tipo atual (Entrada/Saída)
}


def _w() -> int:
    return 66


def _line(ch: str = "=", n: Optional[int] = None) -> str:
    return (ch * (n or _w()))


def _report(msg: str = "") -> None:
    if not _ENABLED:
        return
    _logger.info(msg)


def section(num: str, title: str) -> None:
    if not _ENABLED:
        return
    key = f"section:{num}"
    if key in _state["printed"]:
        return
    _state["printed"].add(key)

    _report("")
    _report(_line("="))
    _report(f"({num}) {title}")
    _report(_line("="))
    _report("")


def subsection(code: str, title: str) -> None:
    if not _ENABLED:
        return
    key = f"sub:{code}:{title}"
    if key in _state["printed"]:
        return
    _state["printed"].add(key)

    _report(f"({code}) " + _line("=", 51))
    _report(title)
    _report("")


def info(msg: str) -> None:
    _report(msg)


def warn(msg: str) -> None:
    _report(f"Aviso: {msg}")


def error(msg: str) -> None:
    _report(f"Erro: {msg}")


def _get_field(fields: Dict[str, Any], *names: str) -> Optional[Any]:
    for n in names:
        if n in fields and fields.get(n) not in (None, ""):
            return fields.get(n)
    return None


def on_step_start(step_name: str, fields: Dict[str, Any]) -> None:
    """
    Chamado pelo trace.step() no início de cada etapa.
    Tenta imprimir secções no estilo do SOMA.py.
    """
    if not _ENABLED:
        return

    sheet = _get_field(fields, "sheet", "ws", "worksheet", "sheet_name")
    if sheet:
        _state["sheet"] = sheet

    row = _get_field(fields, "row", "linha")
    if row is not None:
        _state["current_row"] = row

    tipo = _get_field(fields, "tipo", "type")
    if tipo:
        _state["current_tipo"] = tipo

    # Secções principais (1)(2)(3)(4)(6)
    if step_name == "run.init":
        section("1", "Iniciando o processo")
        if _state.get("sheet"):
            info(f"Selecionando dados na sheet: {_state['sheet']}")
        return

    if step_name == "run.login":
        section("2", "Iniciando o processo de acesso ao SOMA")
        return

    if step_name == "run.preprocess":
        sh = _state.get("sheet") or "-"
        section("3", f"Iniciando o processo de interração com a leitura dos dados da sheet {sh}")
        info("Processando todas as linhas!")
        info("")
        return

    # Ao começar a processar uma linha, inicia a secção (4)
    if step_name in {"row.entrada_saida", "row.saida", "row.entrada"} or step_name.startswith("entradas_saidas."):
        section("4", "Iniciando o processo de interração com a página do SOMA")

    # Subsecções (4.1 ..)
    if step_name == "entradas_saidas.fill_form":
        r = _state.get("current_row")
        t = _state.get("current_tipo") or _get_field(fields, "tipo") or "-"
        subsection("4.1", f"Iniciando o processo de input de dados para a linha {r} e o tipo {t}")
        return

    if step_name in {"entradas_saidas.save_entrada", "entradas_saidas.entrada_transfer_bancaria"}:
        r = _state.get("current_row")
        subsection("4.2", f"Iniciando o processo de inserir pagamento para a linha {r}")
        return

    if step_name == "entradas_saidas.baixa":
        r = _state.get("current_row")
        subsection("4.3", f"Iniciando o processo de baixa do documento para a linha {r}")
        return

    if step_name == "entradas_saidas.search_doc":
        subsection("4.4", "Iniciando o processo pesquisa do nº soma do documento")
        return

    if "caixas" in step_name or "bancos" in step_name:
        section("6", "Iniciando o processo 'Caixas/bancos'")
        return


def on_step_ok(step_name: str, fields: Dict[str, Any], dt_ms: int) -> None:
    """
    Chamado pelo trace.step() quando a etapa termina com sucesso.
    """
    if not _ENABLED:
        return

    # Login / SOMA
    if step_name == "login.fill_credentials":
        info("Preenchimento do utilizador realizado com sucesso")
        info("Preenchimento da senha realizado com sucesso")
        return

    if step_name in {"login.wait_form_disappear", "login.click_login"}:
        info("Click do botão login realizado com sucesso")
        return

    if step_name in {"soma.click_button", "login.open_soma_app"}:
        # Dois pontos diferentes no fluxo; mantém mensagens úteis
        if step_name == "soma.click_button":
            info("Botão 'SOMA' clicado com sucesso!")
        else:
            info("Página carregada com sucesso após clicar no botão 'SOMA'.")
        return

    # Entradas/Saídas – mensagens “humanas”
    if step_name == "entradas_saidas.open_menu":
        info("Clicou na opção 'Entradas/Saídas' com sucesso!")
        return

    if step_name == "entradas_saidas.open_new_form":
        info("Clicou no botão 'Nova Entrada/Saída' com sucesso!")
        return

    if step_name == "entradas_saidas.choose_tipo":
        tipo = _state.get("current_tipo") or _get_field(fields, "tipo") or "-"
        info(f"Selecionado o botão para o processo de '{tipo}'")
        info("")
        return

    if step_name == "entradas_saidas.fill_form":
        info("Campos principais preenchidos com sucesso")
        return

    if step_name == "entradas_saidas.save_entrada":
        info("Documento salvo com sucesso!")
        return

    if step_name == "entradas_saidas.entrada_transfer_bancaria":
        info("Inserir pagamento salvo com sucesso!")
        return

    if step_name == "entradas_saidas.baixa":
        info("Baixa realizada com sucesso!")
        return

    if step_name == "entradas_saidas.search_doc":
        doc = _get_field(fields, "doc", "doc_id")
        if doc:
            info(f"Número do documento extraído: {doc}")
        return


def on_step_fail(step_name: str, fields: Dict[str, Any], dt_ms: int, exc_name: str) -> None:
    """
    Chamado pelo trace.step() quando a etapa falha.
    """
    if not _ENABLED:
        return

    if step_name == "entradas_saidas.back_to_list":
        error("Falha ao voltar para a lista (timeout ao localizar o botão/elemento de retorno).")
        return

    error(f"Falha na etapa '{step_name}' ({exc_name}).")