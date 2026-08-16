from __future__ import annotations

from types import SimpleNamespace

import pytest
from selenium.webdriver.common.by import By

from soma_app.automation.debug_session import GuidedDebugAbort, GuidedDebugSession


class DummyElement:
    def __init__(self, tag="button", text="OK", element_id="", cls="", name="", value=""):
        self.tag_name = tag
        self._attrs = {
            "id": element_id,
            "class": cls,
            "name": name,
            "value": value,
            "outerHTML": f"<{tag} class=\"{cls}\" id=\"{element_id}\">{text}</{tag}>",
        }
        self._text = text

    @property
    def text(self):
        return self._text

    def get_attribute(self, name):
        if name == "outerHTML":
            return self._attrs["outerHTML"]
        return self._attrs.get(name, "")

    def is_displayed(self):
        return True

    def is_enabled(self):
        return True


class DummyDriver:
    def __init__(self):
        self.current_url = "http://example.invalid/popup"
        self.title = "SOMA"
        self.executed = []
        self.requests = []
        self._events = [
            {
                "timestamp": "08:55:10.124",
                "type": "added",
                "tag": "button",
                "id": "",
                "class": "swal2-confirm swal2-styled",
                "text": "OK",
                "url": self.current_url,
            }
        ]
        self.locator_map = {}

    def find_elements(self, by, value):
        self.requests.append((by, value))
        return list(self.locator_map.get((by, value), []))

    def execute_script(self, script, *args):
        self.executed.append(script)
        if "return window.__SOMA_DEBUG_EVENTS" in script:
            return self._events
        if "return document.readyState" in script:
            return "complete"
        if "absoluteXPath" in script and args:
            return "/html/body/div[5]/div/button[1]"
        return "ok"


class DummyActions:
    def __init__(self, driver):
        self.driver = driver
        self.screenshot_calls = []
        self.html_calls = []

    def screenshot(self, name):
        self.screenshot_calls.append(name)
        return f"{name}.png"

    def dump_page_source(self, name):
        self.html_calls.append(name)
        return f"{name}.html"


def test_debug_session_ignores_non_matching_row(monkeypatch):
    monkeypatch.setenv("DEBUG_STEP_MODE", "true")
    monkeypatch.setenv("DEBUG_ROW", "7")

    driver = DummyDriver()
    session = GuidedDebugSession(DummyActions(driver), SimpleNamespace(site_password="secret"))

    called = []
    monkeypatch.setattr("builtins.input", lambda prompt="": called.append(prompt) or "")

    session.checkpoint(
        row=SimpleNamespace(row_number=6),
        stage="SAIDA.TESTE",
        phase="BEFORE",
        action="CLICK",
        element_name="BTN",
        locator=(By.XPATH, "//ignored"),
        instructions=["Não deve pausar."],
    )

    assert called == []


def test_debug_session_commands_and_redaction(monkeypatch, capsys):
    monkeypatch.setenv("DEBUG_STEP_MODE", "true")
    monkeypatch.setenv("DEBUG_ROW", "7")
    monkeypatch.setenv("SITE_PASSWORD", "topsecret")

    driver = DummyDriver()
    driver.locator_map[(By.XPATH, "/html/body/div[5]/div/button[1]")] = [
        DummyElement(tag="button", text="OK", cls="swal2-confirm swal2-styled")
    ]
    driver.locator_map[(By.CSS_SELECTOR, ".swal2-confirm")] = [
        DummyElement(tag="button", text="OK", cls="swal2-confirm swal2-styled")
    ]
    driver.locator_map[(By.TAG_NAME, "button")] = [
        DummyElement(tag="button", text="Salvar", cls="btn-primary")
    ]
    driver.locator_map[(By.TAG_NAME, "iframe")] = []
    driver.locator_map[(By.TAG_NAME, "frame")] = []

    actions = DummyActions(driver)
    session = GuidedDebugSession(actions, SimpleNamespace(site_password="topsecret"))

    commands = iter(["x /html/body/div[5]/div/button[1]", "css .swal2-confirm", "events", "url", "shot", "html", "frames", ""])
    monkeypatch.setattr("builtins.input", lambda prompt="": next(commands))

    session.checkpoint(
        row=SimpleNamespace(row_number=7),
        stage="SAIDA.PAGAMENTO.CONFIRMACAO",
        phase="BEFORE",
        action="WAIT/CHECK",
        element_name="OK_ALERT",
        locator=(By.XPATH, "/html/body/div[5]/div/button[1]"),
        value="topsecret",
        instructions=["Use ENTER para avançar."],
    )

    output = capsys.readouterr().out
    assert "FOUND=1" in output
    assert "swal2-confirm swal2-styled" in output
    assert "events" not in output.lower() or "Eventos DOM/SweetAlert" in output
    assert "topsecret" not in output
    assert actions.screenshot_calls
    assert actions.html_calls
    assert any("MutationObserver" in script for script in driver.executed)


def test_debug_session_q_aborts(monkeypatch):
    monkeypatch.setenv("DEBUG_STEP_MODE", "true")
    monkeypatch.setenv("DEBUG_ROW", "7")

    driver = DummyDriver()
    session = GuidedDebugSession(DummyActions(driver), SimpleNamespace(site_password=""))

    monkeypatch.setattr("builtins.input", lambda prompt="": "q")

    with pytest.raises(GuidedDebugAbort):
        session.checkpoint(
            row=SimpleNamespace(row_number=7),
            stage="SAIDA.TESTE",
            phase="BEFORE",
            action="CLICK",
            element_name="BTN",
            locator=(By.XPATH, "//ignored"),
            instructions=["Teste de abortar."],
        )


def test_debug_step_mode_requires_visible_browser(monkeypatch):
    monkeypatch.setenv("DEBUG_STEP_MODE", "true")
    monkeypatch.setenv("HEADLESS", "true")

    from soma_app.infra.webdriver_factory import create_driver

    with pytest.raises(RuntimeError, match="DEBUG_STEP_MODE=true requer HEADLESS=false"):
        create_driver(settings=SimpleNamespace(headless=True), headless=None, downloads_dir="artifacts/downloads")


def test_create_driver_headless_skips_maximize(monkeypatch):
    from soma_app.infra import webdriver_factory

    calls = {"maximize": 0, "size": 0, "cdp": 0}
    captured = {"timeouts": []}

    class DummyDriver:
        def maximize_window(self):
            calls["maximize"] += 1

        def set_window_size(self, width, height):
            calls["size"] += 1

        def execute_cdp_cmd(self, name, payload):
            calls["cdp"] += 1

    def fake_chrome(*args, **kwargs):
        return DummyDriver()

    monkeypatch.setattr(webdriver_factory.webdriver, "Chrome", fake_chrome)
    monkeypatch.setattr(webdriver_factory, "_build_service", lambda: SimpleNamespace())
    monkeypatch.setattr(webdriver_factory, "_resolve_downloads_dir", lambda settings=None, downloads_dir=None: "artifacts/downloads")
    monkeypatch.setattr(webdriver_factory, "_build_options", lambda headless, downloads_dir: SimpleNamespace())
    monkeypatch.setattr(webdriver_factory, "log_kv", lambda *args, **kwargs: None)
    monkeypatch.setattr(webdriver_factory.RemoteConnection, "get_timeout", lambda: 120)
    monkeypatch.setattr(
        webdriver_factory.RemoteConnection,
        "set_timeout",
        lambda timeout: captured["timeouts"].append(timeout),
    )

    driver = webdriver_factory.create_driver(
        settings=SimpleNamespace(headless=True), headless=None, downloads_dir="artifacts/downloads"
    )

    assert calls["maximize"] == 0
    assert calls["size"] == 0
    assert calls["cdp"] == 1
    assert captured["timeouts"] == [300, 120]
    assert isinstance(driver, DummyDriver)
