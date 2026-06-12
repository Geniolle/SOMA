from __future__ import annotations

from soma_app.infra import report


def _reset_report_state() -> None:
    report._state["printed"] = set()
    report._state["sheet"] = None
    report._state["current_row"] = None
    report._state["current_tipo"] = None
    report._state["current_progress"] = None
    report._state["current_total"] = None


def test_fill_form_subsection_includes_progress(monkeypatch):
    messages: list[str] = []
    _reset_report_state()
    monkeypatch.setattr(report, "_ENABLED", True)
    monkeypatch.setattr(report, "_report", lambda msg="": messages.append(msg))

    report.on_step_start(
        "run.process_row",
        {"row": 3375, "tipo": "Entrada", "progress_current": 1, "progress_total": 4},
    )
    report.on_step_start("entradas_saidas.fill_form", {"row": 3375, "tipo": "Entrada"})

    assert "Iniciando o processo de input de dados para a linha 3375 e o tipo Entrada (1/4)" in messages


def test_report_does_not_duplicate_entradas_saidas_success_messages(monkeypatch):
    messages: list[str] = []
    _reset_report_state()
    monkeypatch.setattr(report, "_ENABLED", True)
    monkeypatch.setattr(report, "_report", lambda msg="": messages.append(msg))

    report.on_step_ok("entradas_saidas.open_menu", {"row": 3375, "tipo": "Entrada"}, 10)

    assert messages == []
