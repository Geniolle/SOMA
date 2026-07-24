from __future__ import annotations

from types import SimpleNamespace

import pytest

from soma_app.workflows import run_soma as run_soma_module


class FakeWebDriverFactory:
    def __init__(self, bundle):
        self.bundle = bundle

    def create(self, settings=None, **kwargs):
        return self.bundle


class BoomApiClient:
    def __init__(self, **kwargs):
        raise RuntimeError("falha ao construir SomaApiClient")


def test_bootstrap_backend_keeps_bundle_reference_when_api_init_fails_after_browser(monkeypatch):
    fake_bundle = SimpleNamespace(a=SimpleNamespace())
    factory = FakeWebDriverFactory(fake_bundle)

    monkeypatch.setattr(run_soma_module, "WebDriverFactory", SimpleNamespace(create=factory.create))
    monkeypatch.setattr(run_soma_module, "unwrap_webdriver", lambda obj: None)
    monkeypatch.setattr(
        run_soma_module,
        "get_chromedriver_info",
        lambda wd: {"version": "", "path": "", "source": "unknown"},
    )
    monkeypatch.setattr(run_soma_module, "get_chrome_version", lambda wd: "")
    monkeypatch.setattr(run_soma_module, "SomaApiClient", BoomApiClient)

    state = run_soma_module.RunState()

    with pytest.raises(RuntimeError, match="falha ao construir SomaApiClient"):
        run_soma_module._bootstrap_backend(
            state,
            settings=None,
            run_id="t1",
            ws="CONTAORDEM",
            headless=True,
            backend_mode="api",
            api_first=True,
            run_caixas_bancos=True,
            api_fallback_selenium=False,
        )

    # A referência ao bundle criado antes da falha precisa sobreviver,
    # senão main() não consegue chamar bundle.quit() no finally e o Chrome fica órfão.
    assert state.bundle is fake_bundle
    assert state.api_client is None
