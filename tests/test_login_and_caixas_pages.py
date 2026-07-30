from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException

from soma_app.automation.pages.caixas_bancos_page import CaixasBancosPage
from soma_app.automation.pages.login_page import LoginPage


class FakeLoginActions:
    def __init__(self, present):
        self.present = set(present)
        self.get_calls: list[str] = []
        self.clicked = []
        self.js_clicked = []
        self.wait_visible_calls = []
        self.wait_invisible_calls = []
        self.driver = SimpleNamespace(
            current_url="https://test.local/",
            title="SOMA",
            window_handles=["main"],
            get=self.get,
            execute_script=lambda *args, **kwargs: None,
            switch_to=SimpleNamespace(window=lambda handle: None),
        )

    def get(self, url: str) -> None:
        self.get_calls.append(url)

    def wait_any_present(self, locators, timeout_seconds=None):
        for loc in locators:
            if loc in self.present:
                return loc
        raise TimeoutException("missing")

    def click_js(self, locator):
        self.js_clicked.append(locator)

    def click(self, locator):
        self.clicked.append(locator)

    def wait_visible(self, locator, timeout_seconds=None):
        self.wait_visible_calls.append(locator)
        if locator in self.present:
            return SimpleNamespace(text="", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("missing")

    def wait_invisible(self, locator, timeout_seconds=None):
        self.wait_invisible_calls.append(locator)
        return None

    def screenshot(self, name):
        return Path(f"{name}.png")

    def wait_dom_ready(self, timeout_seconds=30):
        return None


class FakeCaixasActions:
    def __init__(self):
        self.get_calls: list[str] = []
        self.clicked = []
        self.js_clicked = []
        self.present = set()
        self.home_url = "https://verbodavida.info/IVV/"
        self.driver = SimpleNamespace(
            current_url="https://test.local/IVV/",
            title="SOMA",
            get=self.get,
            execute_script=lambda *args, **kwargs: None,
        )

    def get(self, url: str) -> None:
        self.get_calls.append(url)
        if url in {self.home_url, "https://verbodavida.info/apps/index.php"}:
            self.present.add(("xpath", "//soma"))

    def wait_any_present(self, locators, timeout_seconds=None):
        for loc in locators:
            if loc in self.present:
                return loc
        raise TimeoutException("missing")

    def click_js(self, locator):
        self.js_clicked.append(locator)

    def click(self, locator):
        self.clicked.append(locator)

    def wait_dom_ready(self, timeout_seconds=30):
        return None

    def screenshot(self, name):
        return Path(f"{name}.png")

    def wait_visible(self, locator, timeout_seconds=None):
        if locator in self.present:
            return SimpleNamespace(text="ok", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("missing")


def test_open_soma_app_without_reloading_portal():
    button = ("xpath", "//button[@id='soma']")
    ready = ("xpath", "//div[@id='ready']")
    actions = FakeLoginActions({button, ready})
    settings = SimpleNamespace(site_login_url="https://verbodavida.info/apps/index.php", timeout_seconds=5)
    page = LoginPage(actions, settings)
    page.SOMA_BUTTON_CANDIDATES = [button]
    page.SOMA_READY = ready

    page.open_soma_app(reload_portal=False)

    assert actions.get_calls == []
    assert actions.js_clicked == [button]
    assert actions.wait_visible_calls == [ready]


def test_caixas_open_resets_quickly_when_menu_is_absent():
    menu = ("xpath", "//menu")
    soma = ("xpath", "//soma")
    value = ("xpath", "//value")
    settings = SimpleNamespace(
        site_home_url="https://verbodavida.info/IVV/",
        site_login_url="https://verbodavida.info/apps/index.php",
        timeout_seconds=5,
    )
    actions = FakeCaixasActions()
    page = CaixasBancosPage(actions, settings)
    page.MENU_CAIXAS_BANCOS_CANDIDATES = [menu]
    page.SOMA_BUTTON_CANDIDATES = [soma]
    page.ANY_VALUE_CANDIDATES = [value]

    def click_js(locator):
        actions.js_clicked.append(locator)
        if locator == soma:
            actions.present.update({menu, value})

    actions.click_js = click_js

    page.open()

    assert actions.get_calls[0] == "https://verbodavida.info/apps/index.php"
    assert actions.js_clicked[:2] == [soma, menu]
    assert value in actions.present
