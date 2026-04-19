from __future__ import annotations

from pathlib import Path

import pytest

from soma_app.config.settings import Settings


def _clear_required(monkeypatch: pytest.MonkeyPatch) -> None:
    for key in (
        "GOOGLE_CREDENTIALS_PATH",
        "SPREADSHEET_URL",
        "SITE_USER",
        "SITE_PASSWORD",
    ):
        monkeypatch.delenv(key, raising=False)


def test_settings_from_env_ok(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    cred_file = tmp_path / "credenciais.json"
    cred_file.write_text("{}", encoding="utf-8")

    _clear_required(monkeypatch)
    monkeypatch.setenv("GOOGLE_CREDENTIALS_PATH", str(cred_file))
    monkeypatch.setenv("SPREADSHEET_URL", "https://docs.google.com/spreadsheets/d/test-id/edit")
    monkeypatch.setenv("SITE_USER", "bot@example.com")
    monkeypatch.setenv("SITE_PASSWORD", "secret")
    monkeypatch.setenv("HEADLESS", "true")
    monkeypatch.setenv("TIMEOUT_SECONDS", "35")

    settings = Settings.from_env(env_file=None)

    assert settings.google_credentials_path == cred_file
    assert settings.headless is True
    assert settings.timeout_seconds == 35


def test_settings_missing_required_raises(monkeypatch: pytest.MonkeyPatch):
    _clear_required(monkeypatch)

    with pytest.raises(ValueError):
        Settings.from_env(env_file=None)
