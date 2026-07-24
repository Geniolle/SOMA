from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from selenium.webdriver.common.by import By

import soma_app.automation.interaction_recorder as recorder_module
from soma_app.automation.dom_inventory import DomInventory
from soma_app.automation.interaction_recorder import (
    InteractionRecorder,
    build_workflow_summary,
    consolidate_events,
    consolidate_select2_events,
    dedupe_input_events,
    dedupe_state_events,
    event_to_step,
    sanitize_record_value,
    suggest_wait_condition,
)


class FakeFrame:
    def __init__(self, index: int):
        self.index = index


class FakeSwitchTo:
    def __init__(self, driver: "FakeDriver"):
        self.driver = driver

    def window(self, handle: str) -> None:
        if handle not in self.driver.window_handles:
            raise RuntimeError("window not found")
        self.driver.current_window_handle = handle
        self.driver.frame_stack = []

    def frame(self, frame: FakeFrame) -> None:
        self.driver.frame_stack.append(frame.index)

    def parent_frame(self) -> None:
        if self.driver.frame_stack:
            self.driver.frame_stack.pop()


class FakeDriver:
    def __init__(self):
        self.contexts: dict[tuple[str, tuple[int, ...]], dict[str, object]] = {}
        self.window_handles = ["w1"]
        self.current_window_handle = "w1"
        self.frame_stack: list[int] = []
        self.switch_to = FakeSwitchTo(self)
        self.saved_screenshots: list[str] = []
        self.add_context("w1", (), installed=False)

    @property
    def context_key(self) -> tuple[str, tuple[int, ...]]:
        return self.current_window_handle, tuple(self.frame_stack)

    @property
    def current_context(self) -> dict[str, object]:
        return self.contexts[self.context_key]

    def add_context(
        self,
        handle: str,
        frame_path: tuple[int, ...],
        *,
        installed: bool = False,
        queue: list[dict[str, object]] | None = None,
        state: dict[str, object] | None = None,
        frame_count: int = 0,
    ) -> None:
        self.contexts[(handle, frame_path)] = {
            "installed": installed,
            "queue": list(queue or []),
            "state": {
                "url": f"https://site/{handle}",
                "title": f"Tela {handle}",
                "interactive_count": 1,
                "modal_count": 0,
                "frame_count": frame_count,
                "text_hash": "1",
                "html_hash": "1",
                "alerts_count": 0,
                **(state or {}),
            },
            "frame_count": frame_count,
        }

    def add_window(self, handle: str, *, queue: list[dict[str, object]] | None = None) -> None:
        self.window_handles.append(handle)
        self.add_context(handle, (), installed=False, queue=queue)

    def add_frame(
        self,
        handle: str,
        parent_path: tuple[int, ...],
        index: int,
        *,
        queue: list[dict[str, object]] | None = None,
    ) -> None:
        parent = self.contexts[(handle, parent_path)]
        parent["frame_count"] = max(int(parent.get("frame_count", 0)), index + 1)
        state = dict(parent["state"])
        state["frame_count"] = int(parent["frame_count"])
        parent["state"] = state
        self.add_context(handle, parent_path + (index,), installed=False, queue=queue)

    def execute_script(self, script: str):
        context = self.current_context
        if "return !!(window.__SOMA_INTERACTION_RECORDER__" in script:
            return bool(context["installed"])
        if "window.__SOMA_INTERACTION_RECORDER__ = recorder" in script:
            context["installed"] = True
            return True
        if "rec && rec.flush" in script:
            queue = list(context["queue"])
            context["queue"] = []
            return queue
        if "rec && rec.discard" in script:
            context["queue"] = []
            return True
        if "interactive_count" in script and "modal_count" in script:
            return dict(context["state"])
        return True

    def find_elements(self, by: str, selector: str):
        if by == By.TAG_NAME and selector == "iframe":
            count = int(self.current_context.get("frame_count", 0))
            return [FakeFrame(index) for index in range(count)]
        return []

    def save_screenshot(self, path: str) -> bool:
        self.saved_screenshots.append(path)
        Path(path).write_bytes(b"")
        return True


def _make_recorder(tmp_path: Path) -> InteractionRecorder:
    recorder = InteractionRecorder(
        process_name="processo:teste/inválido",
        root=tmp_path,
        site_recorder=SimpleNamespace(save_page=lambda snapshot: None),
        dom_inventory=DomInventory(),
    )
    recorder._capture_snapshot = lambda *_args, **_kwargs: None  # type: ignore[method-assign]
    return recorder


def test_sanitize_record_value_redacts_inputs_and_checkbox_state():
    redacted = sanitize_record_value("input", value="123456", value_length=6, value_type="text")
    checkbox = sanitize_record_value("change", checked=True, value_type="checkbox")

    assert redacted == {"value": "[redacted]", "value_length": 6, "value_type": "text"}
    assert checkbox == {"checked": True}


def test_dedupe_input_events_keeps_most_recent_event_for_same_field():
    events = [
        {
            "action": "input",
            "timestamp": "2026-07-24T10:00:00+00:00",
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "id": "campo",
        },
        {
            "action": "input",
            "timestamp": "2026-07-24T10:00:00.300000+00:00",
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "id": "campo",
            "value_length": 5,
        },
        {
            "action": "click",
            "timestamp": "2026-07-24T10:00:01+00:00",
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
        },
    ]

    deduped = dedupe_input_events(events)

    assert len(deduped) == 2
    assert deduped[0]["value_length"] == 5
    assert deduped[1]["action"] == "click"


def test_consolidate_select2_events_collapses_sequence():
    events = [
        {
            "action": "click",
            "is_select2": True,
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "css_selector": ".select2-container",
        },
        {
            "action": "select2_input",
            "is_select2": True,
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "css_selector": ".select2-search__field",
        },
        {
            "action": "select2_change",
            "is_select2": True,
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "css_selector": ".select2-results__option",
        },
    ]

    result = consolidate_select2_events(events)

    assert len(result) == 1
    assert result[0]["action"] == "select2_choose"
    assert result[0]["opener_selector"] == ".select2-container"
    assert result[0]["search_selector"] == ".select2-search__field"
    assert result[0]["result_selector"] == ".select2-results__option"


def test_dedupe_state_events_removes_python_and_js_duplicates():
    event = {
        "action": "url_changed",
        "timestamp": "2026-07-24T10:00:00+00:00",
        "window_handle": "w1",
        "iframe_path": "top",
        "before_signature": "a",
        "after_signature": "b",
    }
    duplicate = {**event, "timestamp": "2026-07-24T10:00:00.200000+00:00"}

    result = dedupe_state_events([event, duplicate])

    assert len(result) == 1
    assert result[0]["timestamp"].endswith("200000+00:00")


def test_consolidate_events_combines_input_select2_and_state_noise():
    events = [
        {
            "action": "input",
            "timestamp": "2026-07-24T10:00:00+00:00",
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "id": "campo",
        },
        {
            "action": "input",
            "timestamp": "2026-07-24T10:00:00.200000+00:00",
            "page_url": "https://site",
            "window_handle": "w1",
            "iframe_path": "top",
            "id": "campo",
            "value_length": 7,
        },
        {
            "action": "url_changed",
            "timestamp": "2026-07-24T10:00:01+00:00",
            "window_handle": "w1",
            "iframe_path": "top",
            "before_signature": "a",
            "after_signature": "b",
        },
        {
            "action": "url_changed",
            "timestamp": "2026-07-24T10:00:01.100000+00:00",
            "window_handle": "w1",
            "iframe_path": "top",
            "before_signature": "a",
            "after_signature": "b",
        },
    ]

    result = consolidate_events(events)

    assert len(result) == 2
    assert result[0]["value_length"] == 7
    assert result[1]["action"] == "url_changed"


def test_suggest_wait_condition_detects_relevant_state_changes():
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
        "iframe_path": "top/0",
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


def test_event_to_step_redacts_values_and_sanitizes_sensitive_url():
    event = {
        "action": "change",
        "timestamp": "2026-07-24T10:00:00+00:00",
        "page_url": "https://site/form?token=secret&mode=edit",
        "page_title": "Tela",
        "window_index": 0,
        "window_handle": "w1",
        "iframe_path": "top",
        "tag": "input",
        "type": "text",
        "id": "campo",
        "name": "campo",
        "label": "Campo",
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

    assert step["value"] == "[redacted]"
    assert "secret" not in step["page_url"]
    assert "token=%5Bredacted%5D" in step["page_url"]


def test_workflow_summary_uses_real_windows_and_iframes():
    summary = build_workflow_summary(
        "processo_teste",
        [
            {
                "page_url": "https://site/a?b=2&a=1",
                "window_index": 1,
                "window_handle": "w2",
                "iframe_path": "top/0",
                "field": "Campo",
                "action": "click",
                "label": "Abrir",
                "css_selector": '[id="abrir"]',
                "selector_candidates": [{"unique": True, "score": 90.0}],
                "wait_condition": "dom_changed|modal_visible",
            }
        ],
    )

    assert summary["windows_used"] == ["1:w2"]
    assert summary["iframes_used"] == ["top/0"]
    assert summary["fields_filled"] == ["campo"]
    assert summary["buttons_clicked"] == ["abrir"]
    assert summary["wait_conditions_suggested"] == ["dom_changed", "modal_visible"]


def test_pause_discards_pending_and_paused_events(tmp_path):
    recorder = _make_recorder(tmp_path)
    driver = FakeDriver()
    driver.current_context["queue"] = [{"action": "click", "tag": "button", "text": "Antes"}]

    recorder.pause(driver)
    assert driver.current_context["queue"] == []

    driver.current_context["queue"] = [{"action": "click", "tag": "button", "text": "Durante"}]
    assert recorder.poll(driver) == []
    assert driver.current_context["queue"] == []

    recorder.resume(driver)
    driver.current_context["queue"] = [{"action": "click", "tag": "button", "text": "Depois"}]
    recorder.poll(driver)

    texts = [item.get("text") for item in recorder.raw_events]
    assert "Durante" not in texts
    assert "Depois" in texts


def test_request_stop_flushes_last_pending_event(tmp_path):
    recorder = _make_recorder(tmp_path)
    driver = FakeDriver()
    driver.current_context["queue"] = [
        {"action": "click", "tag": "button", "text": "Último clique"}
    ]

    recorder.request_stop(driver)

    assert recorder.stopped is True
    assert any(item.get("text") == "Último clique" for item in recorder.raw_events)
    assert any(item.get("label") == "stop" for item in recorder.raw_events)


def test_poll_reinjects_after_navigation_and_tracks_new_window(tmp_path):
    recorder = _make_recorder(tmp_path)
    driver = FakeDriver()
    recorder.poll(driver)
    assert driver.contexts[("w1", ())]["installed"] is True

    driver.contexts[("w1", ())]["installed"] = False
    driver.contexts[("w1", ())]["queue"] = [
        {"action": "click", "tag": "button", "text": "Após navegação"}
    ]
    driver.add_window(
        "w2",
        queue=[{"action": "click", "tag": "button", "text": "Nova janela"}],
    )

    recorder.poll(driver)

    assert driver.contexts[("w1", ())]["installed"] is True
    assert driver.contexts[("w2", ())]["installed"] is True
    assert any(item.get("action") == "new_window" and item.get("window_handle") == "w2" for item in recorder.raw_events)
    assert any(item.get("text") == "Nova janela" and item.get("window_index") == 1 for item in recorder.raw_events)


def test_iframe_events_are_validated_while_driver_is_inside_frame(tmp_path, monkeypatch):
    recorder = _make_recorder(tmp_path)
    driver = FakeDriver()
    driver.add_frame(
        "w1",
        (),
        0,
        queue=[{"action": "click", "tag": "button", "id": "inside-frame", "text": "Abrir"}],
    )
    contexts: list[tuple[int, ...]] = []

    def fake_candidates(current_driver, event):
        contexts.append(tuple(current_driver.frame_stack))
        return ([{"strategy": "id", "by": "id", "selector": "inside-frame", "count": 1, "unique": True, "score": 95, "reason": "ok"}], {"selector": "inside-frame", "unique": True, "score": 95})

    monkeypatch.setattr(recorder_module, "build_selector_candidate_payload", fake_candidates)

    recorder.poll(driver)

    assert (0,) in contexts
    frame_event = next(item for item in recorder.raw_events if item.get("id") == "inside-frame")
    assert frame_event["iframe_path"] == "top/0"
    assert frame_event["selector_candidates"][0]["unique"] is True


def test_sensitive_event_fields_are_removed_or_redacted(tmp_path):
    recorder = _make_recorder(tmp_path)
    driver = FakeDriver()
    driver.current_context["queue"] = [
        {
            "action": "alert",
            "tag": "form",
            "text": "Nome completo e valor 500",
            "form_text": "Dados pessoais",
            "message": "Documento 123456789",
            "requested_url": "https://site/callback?token=secret",
            "data_attrs": {"data-testid": "alert", "data-customer-name": "Clayton"},
        }
    ]

    recorder.poll(driver)

    event = next(item for item in recorder.raw_events if item.get("action") == "alert")
    assert event["message"] == "[redacted]"
    assert event["text"] == ""
    assert "form_text" not in event
    assert "secret" not in event["requested_url"]
    assert event["data_attrs"] == {"data-testid": "alert"}


def test_finalize_writes_complete_status_and_safe_filenames(tmp_path):
    recorder = _make_recorder(tmp_path)
    result = recorder.finalize()

    status = json.loads((tmp_path / "record_status.json").read_text(encoding="utf-8"))
    assert status["status"] == "complete"
    assert result["workflow_summary"]["process_name"] == "processo:teste/inválido"
    assert recorder.safe_process_name == "processo_teste_inválido"
    assert (tmp_path / "steps.json").exists()
    assert (tmp_path / "workflow.json").exists()
    assert (tmp_path / "workflow_summary.json").exists()
    assert (tmp_path / "elements_used.json").exists()
    assert (tmp_path / "locator_candidates_used.json").exists()
