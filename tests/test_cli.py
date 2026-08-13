from __future__ import annotations

from soma_app import cli


def test_cli_dispatches_to_audit_command(monkeypatch):
    calls = []

    monkeypatch.setattr("soma_app.workflows.audit_conciliacao.main", lambda: calls.append("audit") or 0)

    assert cli.main(["audit-conciliacao"]) == 0
    assert calls == ["audit"]


def test_cli_defaults_to_run_workflow(monkeypatch):
    calls = []

    monkeypatch.setattr("soma_app.workflows.run_soma.main", lambda: calls.append("run") or 0)

    assert cli.main([]) == 0
    assert calls == ["run"]
