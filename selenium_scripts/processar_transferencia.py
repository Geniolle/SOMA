from __future__ import annotations

import importlib

from selenium_scripts._compat import (
    build_actions,
    build_default_updates,
    load_settings,
    to_row,
    update_sheet_fields,
)


def processar_transferencia(driver, linha, index, sheet):
    settings = load_settings()
    actions = build_actions(driver, settings)

    page_module = importlib.import_module("soma_app.automation.pages.transferencias_page")
    page = page_module.TransferenciasPage(actions, settings)

    row = to_row(index=index, linha=linha, tipo_forcado="Transferência")
    doc_id = str(page.run(row) or "Transferido")

    update_sheet_fields(
        sheet,
        index,
        build_default_updates(doc_soma=doc_id),
    )

    print(f"Transferência finalizada para a linha {index}.")
