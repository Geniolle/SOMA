from __future__ import annotations

import time
from types import SimpleNamespace

from soma_app.tools.site_mapper import RecordConfig, SiteMapperConfig, SiteMapRunner


class FakeDriver:
    def __init__(self) -> None:
        self.current_window_handle = "w1"
        self.window_handles = ["w1"]
        self.calls: list[str] = []


class FakeSession:
    def __init__(self, calls: list[str], *, delay: float = 0.0, should_fail: bool = False) -> None:
        self.calls = calls
        self.delay = delay
        self.should_fail = should_fail
        self.stopped = False

    def install(self, driver) -> None:
        self.calls.append("install")

    def record_checkpoint_event(self, driver, label: str):
        self.calls.append(f"checkpoint_event:{label}")
        if self.delay:
            time.sleep(self.delay)
        if self.should_fail:
            raise RuntimeError("checkpoint failed")
        return {"label": label}

    def capture_checkpoint(self, driver, label: str, *, timeout_seconds: float | None = None):
        self.calls.append(f"checkpoint:{label}")
        if self.delay:
            time.sleep(self.delay)
        if self.should_fail:
            raise RuntimeError("capture failed")
        return {"label": label, "timeout_seconds": timeout_seconds}


def _make_runner(tmp_path, timeout_seconds: float = 0.25) -> SiteMapRunner:
    settings = SimpleNamespace(screenshots_dir=str(tmp_path / "screenshots"), headless=False)
    config = SiteMapperConfig(record=RecordConfig(initial_capture_timeout_seconds=timeout_seconds))
    runner = SiteMapRunner(settings, config=config, run_id="teste", mode="record", headless=False)
    runner._drain_commands = lambda *args, **kwargs: None  # type: ignore[method-assign]
    return runner


def test_record_bootstrap_starts_reader_before_initial_checkpoint(tmp_path):
    runner = _make_runner(tmp_path, timeout_seconds=0.5)
    calls: list[str] = []
    session = FakeSession(calls)
    bundle = SimpleNamespace(driver=FakeDriver())

    runner._start_command_reader = lambda: calls.append("reader")  # type: ignore[method-assign]

    ok = runner._bootstrap_record_mode(bundle, session, "Teste SOMA Codex")

    assert ok is True
    assert calls[:3] == ["reader", "install", "checkpoint_event:record_start"]


def test_record_checkpoint_failure_does_not_block_bootstrap(tmp_path):
    runner = _make_runner(tmp_path, timeout_seconds=0.25)
    calls: list[str] = []
    session = FakeSession(calls, delay=0.4, should_fail=True)
    bundle = SimpleNamespace(driver=FakeDriver())

    runner._start_command_reader = lambda: calls.append("reader")  # type: ignore[method-assign]

    started = time.monotonic()
    ok = runner._bootstrap_record_mode(bundle, session, "Teste SOMA Codex")
    elapsed = time.monotonic() - started

    assert ok is False
    assert elapsed < 1.0
    assert calls[0] == "reader"
    assert "checkpoint_event:record_start" in calls
