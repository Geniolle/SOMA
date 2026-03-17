from __future__ import annotations

import logging
from datetime import datetime
from typing import Any, Dict, List, Tuple

from soma_app.infra.trace import step, log_kv
from soma_app.workflows.process_contaordem import SheetsTable
from soma_app.automation.pages.caixas_bancos_page import CaixasBancosPage

log = logging.getLogger("soma_app.workflows.caixas_bancos")


def _now_pt() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def atualizar_caixas_bancos(sheets: Any, actions: Any, settings: Any) -> None:
    """
    Processo (6) conforme SYS_CAIXA.py:
      - abre menu Caixas/Bancos no site
      - lê valores
      - grava na sheet GERENCIAR CAIXAS (linha 2) nas colunas existentes
    """
    ws_caixas = (getattr(settings, "sheet_caixas", "") or "GERENCIAR CAIXAS").strip()

    print()
    print("==================================================================")
    print("(6) Iniciando o processo 'Caixas/bancos'")
    print("==================================================================")
    print()

    page = CaixasBancosPage(actions, settings)

    with step(log, "caixas.run", ws=ws_caixas):
        page.open()
        values: Dict[str, str] = page.read_values()

        t = SheetsTable(sheets, ws_caixas)
        t.load()

        updates: List[Tuple[int, str, Any]] = []
        row_idx = 2

        # só atualiza colunas que EXISTEM na sheet
        for col_name, val in values.items():
            if t.has_col(col_name):
                updates.append((row_idx, col_name, val))

        if t.has_col("TIMESTAMP"):
            updates.append((row_idx, "TIMESTAMP", _now_pt()))

        if not updates:
            log_kv(
                log,
                logging.ERROR,
                "Nenhuma coluna alvo encontrada na sheet de caixas (ou nenhuma leitura útil).",
                ws=ws_caixas,
                values=list(values.keys()),
            )
            return

        t.batch_update_cells(updates)
        log_kv(log, logging.INFO, "Sheet Caixas atualizada.", ws=ws_caixas, row=row_idx, updated=len(updates))
        print("(6.3) Sheet 'GERENCIAR CAIXAS' atualizada com sucesso.\n")