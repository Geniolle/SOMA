from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from soma_app.automation.dom_inventory import DomInventory
from soma_app.automation.interaction_recorder import (
    InteractionRecorder,
    build_workflow_summary,
    consolidate_events,
    consolidate_select2_events,
    dedupe_input_events,
    event_to_step,
    sanitize_record_value,
    suggest_wait_condition,
)


class DummyDriver:
    def __init__(self, state: dict[str, object] | None = None):
        self.state = state or {}
        self.saved_screenshots: list[str] = []

    def execute_script(self, script: str):
        return dict(self.state)

    def save_screenshot(self, path: str) -> bool:
        self.saved_screenshots.append(path)
        Path(path).write_bytes(b"")
        return True


def _make_recorder(tmp_path: Path) -> InteractionRecorder:
    return InteractionRecorder(
        process_name="processo_teste",
        root=tmp_path,
        site_recorder=SimpleNamespace(save_page=lambda snapshot: None),
        dom_inventory=DomInventory(),
    )


def test_sanitize_record_value_redacts_inputs_and_checkbox_state():
    redacted = sanitize_record_value("input", value="123456", value_length=6, value_type="text")
    checkbox = sanitize_record_value("change", checked=True, value_type="checkbox")

    assert redacted["value"] == "[redacted]"
    assert redacted["value_length"] == 6
    assert redacted["value_type"] == "text"
    assert checkbox == {"checked": True}


def test_dedupe_input_events_keeps_most_recent_event_for_same_field():
    events = [
        {"action": "input", "timestamp": "2026-07-24T10:00:00+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "id": "campo"},
        {"action": "input", "timestamp": "2026-07-24T10:00:00.300000+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "id": "campo", "value_length": 5},
        {"action": "click", "timestamp": "2026-07-24T10:00:01+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top"},
    ]

    deduped = dedupe_input_events(events, max_gap_seconds=0.8)

    assert len(deduped) == 2
    assert deduped[0]["value_length"] == 5
    assert deduped[1]["action"] == "click"


def test_consolidate_select2_events_collapses_select2_sequence():
    events = [
        {"action": "click", "is_select2": True, "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "css_selector": ".select2-container"},
        {"action": "select2_input", "is_select2": True, "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "css_selector": ".select2-search__field", "value_length": 4},
        {"action": "select2_change", "is_select2": True, "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "css_selector": ".select2-results__option", "value_length": 4},
    ]

    consolidated = consolidate_select2_events(events)

    assert len(consolidated) == 1
    step = consolidated[0]
    assert step["action"] == "select2_choose"
    assert step["selected_value"] == "[redacted]"
    assert step["opener_selector"] == ".select2-container"
    assert step["search_selector"] == ".select2-search__field"
    assert step["result_selector"] == ".select2-results__option"


def test_consolidate_events_merges_input_noise():
    events = [
        {"action": "input", "timestamp": "2026-07-24T10:00:00+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "id": "campo"},
        {"action": "input", "timestamp": "2026-07-24T10:00:00.200000+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "id": "campo", "value_length": 7},
        {"action": "select2_input", "timestamp": "2026-07-24T10:00:00.400000+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "css_selector": ".select2-search__field"},
        {"action": "select2_change", "timestamp": "2026-07-24T10:00:00.500000+00:00", "page_url": "https://site", "window_handle": "w1", "iframe_path": "top", "css_selector": ".select2-results__option"},
    ]

    consolidated = consolidate_events(events)

    assert len(consolidated) == 2
    assert consolidated[0]["action"] == "input"
    assert consolidated[0]["value_length"] == 7
    assert consolidated[1]["action"] == "select2_choose"


def test_suggest_wait_condition_detects_url_window_and_modal_changes():
    before = {
        "url": "https://site/a",
        "window_handle": "w1",
        "window_index": 0,
        "iframe_path": "top",
        "interactive_count": 10,
        "modal_count": 0,
        "html_hash": "1",
        "text_hash": "1",
    }
    after = {
        "url": "https://site/b",
        "window_handle": "w2",
        "window_index": 1,
        "iframe_path": "top/modal",
        "interactive_count": 11,
        "modal_count": 1,
        "html_hash": "2",
        "text_hash": "2",
    }

    condition = suggest_wait_condition(before, after)

    assert "url_changed" in condition
    assert "window_changed" in condition
    assert "iframe_changed" in condition
    assert "modal_visible" in condition
    assert "element_count_changed" in condition
    assert "dom_changed" in condition


def test_event_to_step_redacts_values_and_reuses_selector_candidates():
    event = {
        "action": "change",
        "timestamp": "2026-07-24T10:00:00+00:00",
        "page_url": "https://site",
        "page_title": "Tela",
        "window_index": 0,
        "window_handle": "w1",
        "iframe_path": "top",
        "tag": "input",
        "type": "text",
        "id": "campo",
        "name": "campo",
        "class_name": "field",
        "role": "",
        "label": "Campo",
        "aria_label": "",
        "placeholder": "",
        "text": "",
        "css_selector": 'input[id="campo"]',
        "relative_xpath": "//*[@id='campo']",
        "absolute_xpath": "//*[@id='campo']",
        "selector_candidates": [
            {"strategy": "id", "by": "css selector", "selector": '[id="campo"]', "count": 1, "unique": True, "score": 95.0, "reason": "id único"},
        ],
        "selector_recommended": {"strategy": "id", "by": "css selector", "selector": '[id="campo"]', "count": 1, "unique": True, "score": 95.0, "reason": "id único"},
        "value": "[redacted]",
        "value_length": 12,
        "value_type": "text",
    }
    step = event_to_step(
        event,
        step_number=1,
        before_state={"url": "https://site/a", "html_hash": "1", "text_hash": "1"},
        after_state={"url": "https://site/b", "modal_count": 1, "html_hash": "2", "text_hash": "2"},
    )

    assert step["action"] == "change"
    assert step["value"] == "[redacted]"
    assert step["selector_candidates"][0]["selector"] == '[id="campo"]'
    assert step["wait_condition"] == "url_changed|modal_visible|dom_changed"


def test_workflow_summary_groups_pages_windows_and_buttons():
    summary = build_workflow_summary(
        "processo_teste",
        [
            {
                "page_url": "https://site/a?b=2&a=1",
                "window_index": 0,
                "window_handle": "w1",
                "iframe_path": "top",
                "field": "Campo",
                "action": "click",
                "label": "Salvar",
                "css_selector": '[id="salvar"]',
                "relative_xpath": "//*[@id='salvar']",
                "absolute_xpath": "//*[@id='salvar']",
                "selector_candidates": [{"unique": True, "score": 90.0, "reason": "ok"}],
                "wait_condition": "dom_changed",
            }
        ],
    )

    assert summary["process_name"] == "processo_teste"
    assert summary["step_count"] == 1
    assert summary["pages_visited"] == ["https://site/a?a=1&b=2"]
    assert summary["windows_used"] == ["0:w1"]
    assert summary["iframes_used"] == ["top"]
    assert summary["fields_filled"] == ["campo"]
    assert summary["buttons_clicked"] == ["salvar"]
    assert summary["wait_conditions_suggested"]


def test_pause_resume_and_safe_finalization(tmp_path):
    recorder = _make_recorder(tmp_path)
    driver = DummyDriver({"url": "https://site", "title": "Tela", "window_handle": "w1", "iframe_path": "top"})
    recorder._capture_snapshot = lambda *_args, **_kwargs: SimpleNamespace(signature="sig-1", url="https://site")  # type: ignore[method-assign]

    assert recorder.process_command("", driver) == "checkpoint"
    assert recorder.process_command("pause", driver) == "pause"
    assert recorder.paused is True
    assert recorder.process_command("resume", driver) == "resume"
    assert recorder.paused is False
    assert recorder.process_command("q", driver) == "stop"
    assert recorder.stopped is True

    result = recorder.finalize()

    assert result["workflow_summary"]["process_name"] == "processo_teste"
    assert (tmp_path / "steps.json").exists()
    assert (tmp_path / "workflow.json").exists()
    assert (tmp_path / "workflow_summary.json").exists()
    assert (tmp_path / "elements_used.json").exists()
    assert (tmp_path / "locator_candidates_used.json").exists()
