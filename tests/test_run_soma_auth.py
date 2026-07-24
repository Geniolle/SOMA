from __future__ import annotations

import logging

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
