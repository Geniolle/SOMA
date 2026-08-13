from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from soma_app.automation.pages.caixas_bancos_page import CaixasBancosPage


class FakeActions:
    def __init__(self, present=None):
        self.present = set(present or [])
        self.clicked = []
        self.clicked_js = []
        self.driver = SimpleNamespace(current_url="https://example.test", title="SOMA")

    def screenshot(self, name):
        return Path(f"{name}.png")

    def wait_any_present(self, locators, timeout_seconds=None):
        for locator in locators:
            if locator in self.present:
                return locator
        raise TimeoutError("missing")

    def wait_visible(self, locator, timeout_seconds=None):
        raise TimeoutError("missing")

    def scroll_into_view(self, locator):
        return None

    def click_js(self, locator):
        self.clicked_js.append(locator)

    def click(self, locator):
        self.clicked.append(locator)

    def wait_dom_ready(self, timeout_seconds=None):
        return None


def test_read_values_returns_empty_dict_when_cards_are_missing():
    settings = SimpleNamespace(timeout_seconds=1, site_login_url="https://example.test")
    page = CaixasBancosPage(FakeActions(), settings)

    values = page.read_values()

    assert values == {}


def test_open_uses_direct_caixas_link_before_generic_title():
    settings = SimpleNamespace(timeout_seconds=1, site_login_url="https://example.test")
    actions = FakeActions()
    page = CaixasBancosPage(actions, settings)

    page.MENU_CAIXAS_BANCOS_CANDIDATES = [page.MENU_CAIXAS_BANCOS_CANDIDATES[0]]
    actions.present.add(page.MENU_CAIXAS_BANCOS_CANDIDATES[0])

    page._open_menu()

    assert actions.clicked_js == [page.MENU_CAIXAS_BANCOS_CANDIDATES[0]]
