from __future__ import annotations

import importlib

from selenium_scripts._compat import (
    build_actions,
    build_default_updates,
    load_settings,
    to_row,
    update_sheet_fields,
)


def processar_saida(driver, linha, index, sheet):
    settings = load_settings()
    actions = build_actions(driver, settings)

    page_module = importlib.import_module("soma_app.automation.pages.entradas_saidas_page")
    page = page_module.EntradasSaidasPage(actions, settings)

    row = to_row(index=index, linha=linha, tipo_forcado="Saída")
    doc_id = str(page.create_and_get_doc_id(row))

    dados_doc = None
    try:
        dados_doc = str(page.fetch_dados_doc(doc_id))
    except Exception:
        dados_doc = None

    update_sheet_fields(
        sheet,
        index,
        build_default_updates(doc_soma=doc_id, dados_doc=dados_doc),
    )

    print(f"Processamento de saída finalizado para a linha {index}.")
