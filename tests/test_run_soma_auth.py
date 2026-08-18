from __future__ import annotations

import logging
from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException

from soma_app.automation.pages.login_page import LoginPage
from soma_app.workflows.run_soma import _bootstrap_api_session, _should_recover_doc


class FakeApiClient:
    def __init__(self, *, open_error: Exception | None = None):
        self.open_error = open_error
        self.open_calls = 0
        self.used_tokens: list[tuple[str, bool]] = []

    def open_session(self) -> str:
        self.open_calls += 1
        if self.open_error is not None:
            raise self.open_error
        return "opened-token"

    def use_session_token(self, token: str, *, close_on_exit: bool = False) -> None:
        self.used_tokens.append((token, close_on_exit))


class FakeLoginActions:
    def __init__(self, *, present=None):
        self.present = set(present or [])
        self.wait_any_visible_calls: list[tuple] = []
        self.wait_visible_calls: list[tuple] = []
        self.wait_present_calls: list[tuple] = []
        self.clicked_js: list[tuple] = []
        self.types: list[tuple] = []

        class _SwitchTo:
            def window(self, handle):
                return None

        class _Element:
            def __init__(self, outer):
                self.outer = outer
                self.value = ""

            def click(self):
                return None

            def clear(self):
                self.value = ""

            def send_keys(self, text):
                self.value += text
                self.outer.types.append(text)

        self.driver = SimpleNamespace(
            current_url="https://example.invalid/portal",
            title="Portal",
            window_handles=["main"],
            switch_to=_SwitchTo(),
            get=lambda url: None,
            execute_script=lambda *args, **kwargs: None,
        )
        self.element = _Element(self)

    def type(self, locator, text, clear=True):
        self.wait_visible_calls.append((locator, None))
        raise TimeoutException("not visible")

    def wait_any_present(self, locators, timeout_seconds=None):
        for locator in locators:
            if locator in self.present:
                return locator
        raise TimeoutException("missing")

    def wait_any_visible_element(self, locators, timeout_seconds=None, *, log_timeout=True):
        self.wait_any_visible_calls.append((tuple(locators), timeout_seconds, log_timeout))
        for locator in locators:
            if locator in self.present:
                return self.element
        raise TimeoutException("missing")

    def click_js(self, locator):
        self.clicked_js.append(locator)

    def click(self, locator):
        self.clicked_js.append(locator)

    def wait_visible(self, locator, timeout_seconds=None):
        self.wait_visible_calls.append((locator, timeout_seconds))
        raise TimeoutException("not visible")

    def wait_present(self, locator, timeout_seconds=None):
        self.wait_present_calls.append((locator, timeout_seconds))
        if locator in self.present:
            return self.element
        raise TimeoutException("not present")

    def screenshot(self, name):
        raise AssertionError("screenshot não deveria ser chamado neste teste")


def test_login_open_soma_app_falls_back_to_presence_when_ready_is_not_visible(monkeypatch):
    actions = FakeLoginActions()
    settings = SimpleNamespace(site_login_url="https://example.invalid/login", timeout_seconds=2)
    page = LoginPage(actions, settings)

    actions.present.update({page.SOMA_BUTTON_CANDIDATES[0], page.SOMA_READY})
    monkeypatch.setattr("soma_app.automation.pages.login_page.time.sleep", lambda *_args, **_kwargs: None)

    page.open_soma_app()

    assert actions.clicked_js == [page.SOMA_BUTTON_CANDIDATES[0]]
    assert actions.wait_any_visible_calls
    assert actions.wait_present_calls == []


def test_login_debug_credentials_fallback_uses_present_element_when_visible_times_out():
    actions = FakeLoginActions(present=[LoginPage.EMAIL, LoginPage.SENHA])
    settings = SimpleNamespace(site_login_url="https://example.invalid/login", timeout_seconds=2, site_user="user", site_password="pass")
    page = LoginPage(actions, settings)

    page._fill_credentials(debug=True)

    assert actions.types == ["user", "pass"]
    assert actions.wait_visible_calls == [(LoginPage.EMAIL, None), (LoginPage.SENHA, None)]
    assert actions.wait_present_calls == [(LoginPage.EMAIL, 30), (LoginPage.SENHA, 30)]


def test_bootstrap_prefers_existing_token_by_default():
    client = FakeApiClient()

    _bootstrap_api_session(
        client,
        session_token="token-123",
        has_api_credentials=True,
        auth_preference="token_first",
        close_on_exit=False,
    )

    assert client.open_calls == 0
    assert client.used_tokens == [("token-123", False)]


def test_bootstrap_falls_back_to_token_when_login_first_fails(caplog):
    client = FakeApiClient(open_error=RuntimeError("boom"))

    with caplog.at_level(logging.WARNING):
        _bootstrap_api_session(
            client,
            session_token="token-123",
            has_api_credentials=True,
            auth_preference="login_first",
            close_on_exit=True,
        )

    assert client.open_calls == 1
    assert client.used_tokens == [("token-123", True)]
    assert "Usando SOMA_SESSION_TOKEN como fallback." in caplog.text


def test_recover_doc_requires_pending_doc_status():
    assert _should_recover_doc({"DOC. SOMA": "EM ERRO", "STATUS": "PENDENTE_DOC"}) is True
    assert _should_recover_doc({"DOC. SOMA": "EM ERRO", "STATUS": "ERRO"}) is False
    assert _should_recover_doc({"DOC. SOMA": "EM ERRO", "STATUS": ""}) is False
    # DOC. SOMA pode ter sido sobrescrito (lock/unlock) entre o run que criou o
    # documento e este reprocessamento; STATUS=PENDENTE_DOC sozinho já deve
    # disparar a recuperação em vez de recriar o documento.
    assert _should_recover_doc({"DOC. SOMA": "", "STATUS": "PENDENTE_DOC"}) is True
