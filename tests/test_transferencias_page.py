from __future__ import annotations

from types import SimpleNamespace

from soma_app.automation.pages.transferencias_page import TransferenciasPage
from soma_app.domain.models import ContaOrdemRow


class FakeRowElement:
    def __init__(self, text: str):
        self.text = text


class FakeDriver:
    def __init__(self, row_texts: list[str]):
        self._row_texts = row_texts
        self.current_url = "http://example.invalid"

    def find_elements(self, by, value):
        if by == "css selector" and value == "table tbody tr":
            return [FakeRowElement(text) for text in self._row_texts]
        return []

    def execute_script(self, *args, **kwargs):
        return None


class FakeActions:
    def __init__(self, row_texts: list[str]):
        self.driver = FakeDriver(row_texts)
        self.typed: list[tuple] = []
        self.clicked: list[tuple] = []

    def type(self, locator, value, clear=True):
        self.typed.append((locator, value, clear))

    def click_js(self, locator):
        self.clicked.append(locator)

    def wait_dom_ready(self, timeout_seconds=15):
        return None

    def exists(self, locator, timeout_seconds=None):
        return False

    def wait_invisible(self, locator, timeout_seconds=None):
        return None

    def wait_visible(self, locator, timeout_seconds=None):
        return SimpleNamespace(clear=lambda: None, send_keys=lambda *_args, **_kwargs: None)

    def wait_any_present(self, locators, timeout_seconds=None):
        return locators[0]

    def wait_present(self, locator, timeout_seconds=None):
        return SimpleNamespace()

    def set_debug_context(self, context):
        return None

    def screenshot(self, name):
        return SimpleNamespace()


def _build_row() -> ContaOrdemRow:
    raw = {
        "TIPO": "Transferencia",
        "CAIXA SAIDA": "CAIXA DIARIO",
        "CAIXA": "CAIXA DESTINO",
        "DATA MOV.": "31/10/2026",
        "IMPORTANCIA": "1,00",
        "DESCRICAO SOMA": "TRANSFERENCIA ENTRE CAIXAS N001",
        "DOC. SOMA": "",
        "STATUS": "",
    }
    return ContaOrdemRow.from_table_row(3, raw)


def test_transferencia_audit_detects_duplicate_row_before_new_form():
    actions = FakeActions([
        "31/10/2026 | CAIXA DIARIO | CAIXA DESTINO | 1,00 | TRANSFERENCIA ENTRE CAIXAS N001",
        "outra linha sem match",
    ])
    page = TransferenciasPage(actions, SimpleNamespace(site_home_url="", timeout_seconds=20))
    row = _build_row()

    duplicate_row = page._audit_existing_transfer_before_new_v1(row)

    assert "CAIXA DIARIO" in duplicate_row
    assert "CAIXA DESTINO" in duplicate_row


def test_transferencia_row_match_requires_origin_destination_and_value():
    actions = FakeActions(["linha qualquer"])
    page = TransferenciasPage(actions, SimpleNamespace(site_home_url="", timeout_seconds=20))
    row = _build_row()

    assert page._row_matches_transfer_v1("CAIXA DIARIO CAIXA DESTINO 31/10/2026 1,00", row) is True
    assert page._row_matches_transfer_v1("CAIXA DIARIO CAIXA DESTINO 31/10/2026 2,00", row) is False
