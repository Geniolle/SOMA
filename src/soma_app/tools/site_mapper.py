from __future__ import annotations

import argparse
import csv
import html
import json
import logging
import os
import queue
import threading
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping

from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, WebDriverException

from soma_app.automation.actions import ActionConfig, Actions
from soma_app.automation.dom_inventory import (
    CaptureTracker,
    DomInventory,
    PageSnapshot,
    SelectorCandidate,
    is_dangerous_selector,
    is_dangerous_text,
    normalize_text,
    normalize_url,
    sanitize_json_value,
    sanitize_text,
)
from soma_app.automation.interaction_recorder import InteractionRecorder, collect_page_state, page_state_signature
from soma_app.automation.pages.login_page import LoginPage
from soma_app.config.locators import _coerce_locator, _coerce_locator_list, load_page_locator_config
from soma_app.config.settings import Settings
from soma_app.infra.log_config import configure_logging, ensure_artifacts_dirs
from soma_app.infra.trace import log_kv, new_run_id, step
from soma_app.infra.webdriver_factory import WebDriverFactory

log = logging.getLogger("soma_app.tools.site_mapper")

DEFAULT_CONFIG_PATH = Path(__file__).resolve().parents[1] / "config" / "site_mapper.json"


@dataclass
class ManualConfig:
    poll_seconds: float = 1.0
    stability_seconds: float = 1.0
    capture_hidden: bool = False


@dataclass
class ControlledConfig:
    allowed_url_patterns: list[str] = field(default_factory=list)
    allowed_menu_texts: list[str] = field(default_factory=list)
    allowed_click_selectors: list[str] = field(default_factory=list)
    blocked_texts: list[str] = field(default_factory=list)
    blocked_selectors: list[str] = field(default_factory=list)
    blocked_form_selectors: list[str] = field(default_factory=list)
    max_pages: int = 50
    max_depth: int = 3
    timeout_seconds: int = 30
    stability_seconds: float = 1.0
    capture_hidden: bool = False
    max_frames: int = 1


@dataclass
class RecordConfig:
    poll_seconds: float = 0.5
    stability_seconds: float = 1.0
    capture_hidden: bool = False
    initial_capture_timeout_seconds: float = 20.0


@dataclass
class SiteMapperConfig:
    manual: ManualConfig = field(default_factory=ManualConfig)
    controlled: ControlledConfig = field(default_factory=ControlledConfig)
    record: RecordConfig = field(default_factory=RecordConfig)


def _load_settings() -> Settings:
    project_root = Path(__file__).resolve().parents[3]

    env_file = (os.getenv("ENV_FILE") or "").strip()
    if env_file:
        env_path = Path(env_file)
        if not env_path.is_absolute():
            cwd_candidate = Path.cwd() / env_path
            if cwd_candidate.exists():
                return Settings.from_env(env_file=str(cwd_candidate))

            project_candidate = project_root / env_path
            if project_candidate.exists():
                return Settings.from_env(env_file=str(project_candidate))
        elif env_path.exists():
            return Settings.from_env(env_file=str(env_path))

    default_env = project_root / "deploy" / ".env"
    if default_env.exists():
        return Settings.from_env(env_file=str(default_env))

    return Settings.from_env(env_file=None)


def _read_config_path(raw_path: str | None) -> Path:
    if not raw_path:
        return DEFAULT_CONFIG_PATH
    path = Path(raw_path)
    if path.is_absolute():
        return path
    if path.exists():
        return path.resolve()
    candidate = (Path.cwd() / path).resolve()
    if candidate.exists():
        return candidate
    return (DEFAULT_CONFIG_PATH.parent / path.name).resolve()


def load_site_mapper_config(path: str | None = None) -> SiteMapperConfig:
    cfg_path = _read_config_path(path)
    if not cfg_path.exists():
        return SiteMapperConfig()

    data = json.loads(cfg_path.read_text(encoding="utf-8"))
    manual_cfg = data.get("manual", {}) if isinstance(data, dict) else {}
    controlled_cfg = data.get("controlled", {}) if isinstance(data, dict) else {}
    record_cfg = data.get("record", {}) if isinstance(data, dict) else {}

    manual = ManualConfig(
        poll_seconds=float(manual_cfg.get("poll_seconds", 1.0)),
        stability_seconds=float(manual_cfg.get("stability_seconds", 1.0)),
        capture_hidden=bool(manual_cfg.get("capture_hidden", False)),
    )
    controlled = ControlledConfig(
        allowed_url_patterns=list(controlled_cfg.get("allowed_url_patterns", []) or []),
        allowed_menu_texts=list(controlled_cfg.get("allowed_menu_texts", []) or []),
        allowed_click_selectors=list(controlled_cfg.get("allowed_click_selectors", []) or []),
        blocked_texts=list(controlled_cfg.get("blocked_texts", []) or []),
        blocked_selectors=list(controlled_cfg.get("blocked_selectors", []) or []),
        blocked_form_selectors=list(controlled_cfg.get("blocked_form_selectors", []) or []),
        max_pages=int(controlled_cfg.get("max_pages", 50)),
        max_depth=int(controlled_cfg.get("max_depth", 3)),
        timeout_seconds=int(controlled_cfg.get("timeout_seconds", 30)),
        stability_seconds=float(controlled_cfg.get("stability_seconds", 1.0)),
        capture_hidden=bool(controlled_cfg.get("capture_hidden", False)),
        max_frames=int(controlled_cfg.get("max_frames", 1)),
    )
    record = RecordConfig(
        poll_seconds=float(record_cfg.get("poll_seconds", 0.5)),
        stability_seconds=float(record_cfg.get("stability_seconds", 1.0)),
        capture_hidden=bool(record_cfg.get("capture_hidden", False)),
        initial_capture_timeout_seconds=float(record_cfg.get("initial_capture_timeout_seconds", 20.0)),
    )
    return SiteMapperConfig(manual=manual, controlled=controlled, record=record)


def _ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def _safe_filename(value: str) -> str:
    cleaned = []
    for ch in value:
        if ch.isalnum() or ch in {"-", "_", "."}:
            cleaned.append(ch)
        else:
            cleaned.append("_")
    out = "".join(cleaned).strip("._")
    return out or "page"


def _json_dump(path: Path, payload: Any) -> None:
    path.write_text(json.dumps(sanitize_json_value(payload), ensure_ascii=False, indent=2), encoding="utf-8")


def _csv_dump(path: Path, rows: Iterable[Mapping[str, Any]]) -> None:
    rows = list(rows)
    if not rows:
        path.write_text("", encoding="utf-8")
        return

    headers = sorted({key for row in rows for key in row.keys()})
    with path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({key: sanitize_json_value(row.get(key)) for key in headers})


def _hash_text(value: str) -> str:
    import hashlib

    return hashlib.sha256((value or "").encode("utf-8")).hexdigest()


def _dom_root_root(settings: Settings, run_id: str) -> Path:
    artifacts_root = Path(settings.screenshots_dir).resolve().parent
    return artifacts_root / "site-map" / run_id


def _read_page_state(driver: Any) -> dict[str, Any]:
    script = """
    const main = document.querySelector('main, [role="main"], #content, .content, .main') || document.body || document.documentElement;
    const text = main ? (main.innerText || main.textContent || '') : '';
    const html = main ? (main.innerHTML || '') : '';
    const interactive = Array.from(document.querySelectorAll('input, select, textarea, button, a, [role], [contenteditable]')).map(el => {
      const rect = el.getBoundingClientRect();
      const label = el.getAttribute('aria-label') || '';
      const text = (el.innerText || el.textContent || '').trim();
      return [el.tagName.toLowerCase(), el.getAttribute('id') || '', el.getAttribute('name') || '', el.getAttribute('type') || '', label, text, Math.round(rect.x || 0), Math.round(rect.y || 0)].join(':');
    }).join('|');
    return {
      url: location.href,
      title: document.title || '',
      handle: window.name || '',
      body_html: document.body ? document.body.innerHTML || '' : '',
      main_html: html,
      interactive: interactive,
      modal_count: document.querySelectorAll('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container').length,
      iframe_count: document.querySelectorAll('iframe').length,
    };
    """
    return driver.execute_script(script) or {}


class SiteMapRecorder:
    def __init__(self, root: Path, *, config: SiteMapperConfig, locators_path: str | None = None):
        self.root = _ensure_dir(root)
        self.pages_dir = _ensure_dir(self.root / "pages")
        self.elements_dir = _ensure_dir(self.root / "elements")
        self.screenshots_dir = _ensure_dir(self.root / "screenshots")
        self.dom_dir = _ensure_dir(self.root / "dom")
        self.config = config
        self.locators_path = locators_path
        self.pages: list[dict[str, Any]] = []
        self.elements: list[dict[str, Any]] = []
        self.locator_candidates: list[dict[str, Any]] = []
        self.locator_report: list[dict[str, Any]] = []
        self._tracker = CaptureTracker()

    def save_page(self, snapshot: PageSnapshot) -> bool:
        if not self._tracker.should_capture(snapshot.signature):
            return False

        page_slug = _safe_filename(snapshot.page_id)
        html_path = self.dom_dir / f"{page_slug}.html"
        json_path = self.dom_dir / f"{page_slug}.json"
        screenshot_path = Path(snapshot.screenshot_path) if snapshot.screenshot_path else self.screenshots_dir / f"{page_slug}.png"

        if snapshot.dom_html_path:
            html_path = Path(snapshot.dom_html_path)

        self.pages.append(snapshot.to_dict())

        for element in snapshot.elements:
            element_dict = element.to_dict()
            element_dict["page_id"] = snapshot.page_id
            element_dict["page_signature"] = snapshot.signature
            element_dict["page_url"] = snapshot.url
            element_dict["page_title"] = snapshot.title
            element_dict["screenshot_path"] = str(screenshot_path)
            self.elements.append(element_dict)

        locator_candidates = []
        for element in snapshot.elements:
            best = None
            if element.selector_candidates:
                best = max(element.selector_candidates, key=lambda item: (item.unique, item.score, -item.found_count))
            locator_candidates.append(
                {
                    "page_id": snapshot.page_id,
                    "element_id": element.element_id,
                    "tag": element.tag,
                    "label": element.label,
                    "text": element.text,
                    "best": asdict(best) if best is not None else None,
                    "candidates": [asdict(item) for item in element.selector_candidates],
                    "selector_confidence": element.selector_confidence,
                    "url": element.url,
                    "window_index": element.window_index,
                    "iframe_path": element.iframe_path,
                }
            )
        self.locator_candidates.extend(locator_candidates)

        _json_dump(json_path, snapshot.to_dict())
        html_content = snapshot.dom_html or ""
        if not html_content.strip():
            html_content = "\n".join(
                [
                    "<!doctype html>",
                    "<html><head><meta charset='utf-8'><title>sanitized snapshot</title></head><body>",
                    f"<pre>{html.escape(sanitize_text(snapshot.title))}</pre>",
                    f"<pre>{html.escape(sanitize_text(snapshot.url))}</pre>",
                    html.escape(snapshot.signature),
                    "</body></html>",
                ]
            )
        html_path.write_text(html_content, encoding="utf-8")

        for frame in snapshot.frames:
            if not frame.html:
                continue
            frame_path = self.dom_dir / f"{_safe_filename(frame.frame_path)}.html"
            frame_path.write_text(frame.html, encoding="utf-8")
            frame.html_path = str(frame_path)
        return True

    def finalize(self, *, run_id: str, settings: Settings) -> dict[str, Any]:
        elements_csv = self.root / "elements.csv"
        pages_json = self.root / "pages.json"
        elements_json = self.root / "elements.json"
        locator_candidates_json = self.root / "locator_candidates.json"
        summary_json = self.root / "summary.json"
        locator_report_json = self.root / "locator_report.json"

        _json_dump(pages_json, self.pages)
        _json_dump(elements_json, self.elements)
        _json_dump(locator_candidates_json, {"run_id": run_id, "pages": self.locator_candidates})
        _csv_dump(elements_csv, self.elements)

        locator_report = self._build_locator_report()
        self.locator_report = locator_report
        _json_dump(locator_report_json, locator_report)

        summary = {
            "run_id": run_id,
            "captured_pages": len(self.pages),
            "captured_elements": len(self.elements),
            "unique_signatures": len({page.get("signature") for page in self.pages}),
            "screenshots_dir": str(self.screenshots_dir),
            "dom_dir": str(self.dom_dir),
            "pages_json": str(pages_json),
            "elements_json": str(elements_json),
            "elements_csv": str(elements_csv),
            "locator_candidates_json": str(locator_candidates_json),
            "locator_report_json": str(locator_report_json),
            "locator_report": locator_report,
            "config": {
                "manual": asdict(self.config.manual),
                "controlled": asdict(self.config.controlled),
            },
        }
        _json_dump(summary_json, summary)
        return summary

    def _build_locator_report(self) -> list[dict[str, Any]]:
        if not self.pages:
            return []

        page_by_signature: dict[str, list[dict[str, Any]]] = {}
        for page in self.pages:
            page_by_signature.setdefault(page.get("signature", ""), []).append(page)

        report: list[dict[str, Any]] = []
        try:
            locators_root = load_page_locator_config(None, "common")
            _ = locators_root
        except Exception:
            pass

        if self.locators_path:
            settings = type("_Tmp", (), {"locators_path": self.locators_path})()
        else:
            settings = None

        # Compara as secções conhecidas do JSON oficial com a captura.
        page_sections = []
        try:
            raw = load_page_locator_config(settings, "login")
            if raw is not None:
                page_sections.append("login")
        except Exception:
            pass

        for page_name in ("login", "entradas_saidas", "transferencias", "caixas_bancos"):
            try:
                cfg = load_page_locator_config(settings, page_name)
            except Exception:
                cfg = {}
            if not cfg:
                continue

            page_entry = {
                "page_name": page_name,
                "locators": [],
                "found_unique": 0,
                "found_duplicated": 0,
                "not_found": 0,
                "fragile": 0,
                "possible_better_selector": 0,
            }

            for name, raw_locator in cfg.items():
                if name == "common":
                    continue
                locators = []
                if isinstance(raw_locator, list):
                    locators = _coerce_locator_list(raw_locator, [])
                else:
                    loc = _coerce_locator(raw_locator, None)
                    if loc:
                        locators = [loc]

                locator_report_item = {
                    "name": name,
                    "status": "not_found",
                    "count": 0,
                    "unique": False,
                    "selector": None,
                    "reason": "",
                }

                best_match = None
                best_count = 0
                best_unique = False

                for page in self.pages:
                    page_elements = [item for item in self.elements if item.get("page_signature") == page.get("signature")]
                    counts = 0
                    matched_element = None
                    for candidate in locators:
                        by, selector = candidate
                        # Reavalia por heurística a partir dos elementos capturados.
                        for element in page_elements:
                            if by == "css selector":
                                if selector and selector in (element.get("css_selector") or ""):
                                    counts += 1
                                    matched_element = element
                            elif by == "xpath":
                                if selector and selector in (element.get("relative_xpath") or ""):
                                    counts += 1
                                    matched_element = element
                            elif by == "id" and selector and selector == element.get("id"):
                                counts += 1
                                matched_element = element
                            elif by == "name" and selector and selector == element.get("name"):
                                counts += 1
                                matched_element = element
                    if counts > best_count:
                        best_count = counts
                        best_match = matched_element
                        best_unique = counts == 1

                if best_count == 1:
                    locator_report_item["status"] = "found_unique"
                    page_entry["found_unique"] += 1
                    if best_match and float(best_match.get("selector_confidence") or 0.0) < 60:
                        locator_report_item["status"] = "fragile"
                        page_entry["fragile"] += 1
                    elif best_match and (best_match.get("absolute_xpath") or "").startswith("/"):
                        locator_report_item["status"] = "fragile"
                        page_entry["fragile"] += 1
                elif best_count > 1:
                    locator_report_item["status"] = "duplicated"
                    page_entry["found_duplicated"] += 1
                else:
                    page_entry["not_found"] += 1

                locator_report_item["count"] = best_count
                locator_report_item["unique"] = best_unique
                if best_match:
                    locator_report_item["selector"] = best_match.get("css_selector") or best_match.get("relative_xpath")
                    locator_report_item["reason"] = "matched captured element"
                page_entry["locators"].append(locator_report_item)

            report.append(page_entry)

        return report


def _allowed_selector_match(selector: str, allowed_selectors: Iterable[str]) -> bool:
    selector_norm = normalize_text(selector)
    for allowed in allowed_selectors:
        allowed_norm = normalize_text(allowed)
        if not allowed_norm:
            continue
        if allowed_norm == selector_norm or allowed_norm in selector_norm or selector_norm in allowed_norm:
            return True
    return False


def _allowed_text_match(text: str, allowed_texts: Iterable[str]) -> bool:
    text_norm = normalize_text(text)
    for allowed in allowed_texts:
        allowed_norm = normalize_text(allowed)
        if allowed_norm and (allowed_norm in text_norm or text_norm in allowed_norm):
            return True
    return False


def _blocked_any(text: str, selector: str, blocked_texts: Iterable[str], blocked_selectors: Iterable[str]) -> bool:
    if is_dangerous_text(text) or is_dangerous_selector(selector):
        return True
    if _allowed_selector_match(selector, blocked_selectors):
        return True
    if _allowed_text_match(text, blocked_texts):
        return True
    return False


class SiteMapRunner:
    def __init__(self, settings: Settings, *, config: SiteMapperConfig, run_id: str, mode: str, headless: bool | None = None):
        self.settings = settings
        self.config = config
        self.run_id = run_id
        self.mode = mode
        self.headless = settings.headless if headless is None else bool(headless)
        self.root = _dom_root_root(settings, run_id)
        self.recorder = SiteMapRecorder(self.root, config=config, locators_path=str(DEFAULT_CONFIG_PATH.parent / "locators.json"))
        self.inventory = DomInventory(
            capture_hidden=config.manual.capture_hidden or config.controlled.capture_hidden,
            max_frame_depth=config.controlled.max_frames,
            logger=log,
        )
        self._stop_event = threading.Event()
        self._manual_capture_event = threading.Event()
        self._seen_windows: set[str] = set()
        self._command_queue: queue.Queue[str] = queue.Queue()
        self.record_session: InteractionRecorder | None = None

    def _build_bundle(self):
        bundle = WebDriverFactory.create(self.settings, headless=self.headless)
        if bundle.a is None:
            bundle.a = Actions(
                bundle.driver,
                ActionConfig(
                    timeout_seconds=int(getattr(self.settings, "timeout_seconds", 20) or 20),
                    screenshots_dir=self.recorder.screenshots_dir,
                ),
            )
        return bundle

    def _login(self, bundle: Any) -> None:
        login = LoginPage(bundle.a, self.settings)
        login.login()

    def _start_command_reader(self) -> threading.Thread:
        def _reader() -> None:
            while not self._stop_event.is_set():
                try:
                    line = input()
                except EOFError:
                    self._stop_event.set()
                    return
                self._command_queue.put(line)

        thread = threading.Thread(target=_reader, daemon=True)
        thread.start()
        return thread

    def _capture_record_checkpoint(self, bundle: Any, session: InteractionRecorder, label: str) -> bool:
        timeout_seconds = max(0.25, float(self.config.record.initial_capture_timeout_seconds))
        log_kv(log, "[record] A iniciar checkpoint inicial", level=logging.INFO, run_id=self.run_id, timeout_seconds=timeout_seconds)
        done = threading.Event()
        result: dict[str, Any] = {}

        def _worker() -> None:
            try:
                result["payload"] = session.capture_checkpoint(
                    bundle.driver,
                    label,
                    timeout_seconds=timeout_seconds,
                )
            except Exception as exc:
                result["error"] = exc
            finally:
                done.set()

        thread = threading.Thread(target=_worker, daemon=True, name=f"soma-record-{label}")
        thread.start()

        deadline = time.monotonic() + timeout_seconds
        poll_interval = min(0.25, max(0.05, timeout_seconds / 10.0))
        while not done.wait(poll_interval):
            self._drain_commands(bundle, session)
            if self._stop_event.is_set() or session.stopped:
                log_kv(log, "[record] Checkpoint inicial interrompido por comando do utilizador.", level=logging.INFO, run_id=self.run_id)
                return False
            if time.monotonic() >= deadline:
                log_kv(log, "[record] Checkpoint inicial excedeu o tempo limite.", level=logging.WARNING, run_id=self.run_id, timeout_seconds=timeout_seconds)
                return False

        if "error" in result:
            log_kv(log, "[record] Checkpoint inicial falhou.", level=logging.WARNING, run_id=self.run_id, err=str(result["error"]))
            return False

        log_kv(log, "[record] Checkpoint inicial concluído.", level=logging.INFO, run_id=self.run_id)
        return True

    def _bootstrap_record_mode(self, bundle: Any, session: InteractionRecorder, process_name: str) -> bool:
        log_kv(log, "[record] A instalar recorder", level=logging.INFO, run_id=self.run_id, process=process_name)
        self._start_command_reader()
        log_kv(log, "[record] Leitor de comandos iniciado", level=logging.INFO, run_id=self.run_id)
        log.info("Modo record ativo. Comandos: Enter=checkpoint, mark=marcador, pause=pausa, resume=retoma, q=finalizar.")

        try:
            session.install(bundle.driver)
            log_kv(log, "[record] Recorder instalado", level=logging.INFO, run_id=self.run_id)
        except Exception as exc:
            log_kv(log, "[record] Falha ao instalar recorder.", level=logging.WARNING, run_id=self.run_id, err=str(exc))

        checkpoint_ok = self._capture_record_checkpoint(bundle, session, "record_start")
        if not checkpoint_ok:
            log_kv(
                log,
                "[record] A gravação continua sem checkpoint inicial concluído.",
                level=logging.WARNING,
                run_id=self.run_id,
            )
        return checkpoint_ok

    def _drain_commands(self, bundle: Any, recorder: InteractionRecorder) -> None:
        while True:
            try:
                command = self._command_queue.get_nowait()
            except queue.Empty:
                return
            result = recorder.process_command(command, bundle.driver)
            if result == "checkpoint":
                log_kv(log, "Checkpoint manual registado.", level=logging.INFO, run_id=self.run_id)
            elif result == "mark":
                log_kv(log, "Marcador registado.", level=logging.INFO, run_id=self.run_id)
            elif result == "pause":
                log_kv(log, "Gravação pausada.", level=logging.INFO, run_id=self.run_id)
            elif result == "resume":
                log_kv(log, "Gravação retomada.", level=logging.INFO, run_id=self.run_id)
            elif result == "stop":
                self._stop_event.set()
                return

    def _capture_current(self, bundle: Any, *, reason: str) -> PageSnapshot | None:
        try:
            page_id = f"{self.run_id}_{len(self.recorder.pages) + 1:04d}"
            screenshot_path = str(self.recorder.screenshots_dir / f"{_safe_filename(page_id)}.png")
            try:
                bundle.a.screenshot(page_id)
            except Exception:
                pass
            snapshot = self.inventory.capture(
                bundle.driver,
                page_id=page_id,
                screenshot_path=screenshot_path,
                dom_dir=str(self.recorder.dom_dir),
            )
            captured = self.recorder.save_page(snapshot)
            if captured:
                log_kv(
                    log,
                    "Página capturada.",
                    level=logging.INFO,
                    run_id=self.run_id,
                    reason=reason,
                    url=snapshot.url,
                    title=snapshot.title,
                    signature=snapshot.signature,
                    elements=snapshot.element_count,
                )
                return snapshot
            return None
        except (InvalidSessionIdException, WebDriverException) as exc:
            log_kv(log, "Falha ao capturar página.", level=logging.ERROR, run_id=self.run_id, reason=reason, err=str(exc))
            return None

    def _manual_input_loop(self) -> None:
        try:
            while not self._stop_event.is_set():
                try:
                    line = input()
                except EOFError:
                    self._stop_event.set()
                    return
                text = line.strip().lower()
                if text in {"q", "quit", "exit"}:
                    self._stop_event.set()
                    return
                self._manual_capture_event.set()
        except Exception:
            self._stop_event.set()

    def _page_state_signature(self, bundle: Any) -> str:
        try:
            state = _read_page_state(bundle.driver)
            payload = "|".join(
                [
                    normalize_url(state.get("url", "")),
                    sanitize_text(state.get("title", "")),
                    _hash_text(state.get("main_html", "")),
                    _hash_text(state.get("interactive", "")),
                    str(state.get("modal_count", 0)),
                    str(state.get("iframe_count", 0)),
                    getattr(bundle.driver, "current_window_handle", ""),
                ]
            )
            return _hash_text(payload)
        except Exception:
            return ""

    def run_manual(self, bundle: Any) -> int:
        self._login(bundle)
        self._capture_current(bundle, reason="login")

        prompt_thread = threading.Thread(target=self._manual_input_loop, daemon=True)
        prompt_thread.start()

        last_state = ""
        log.info("Modo manual ativo. Prima Enter para capturar a página atual, ou escreva q e Enter para terminar.")
        try:
            while not self._stop_event.is_set():
                try:
                    handles = list(bundle.driver.window_handles)
                except Exception:
                    handles = []

                current_handles = [handle for handle in handles if handle not in self._seen_windows]
                for handle in current_handles:
                    try:
                        bundle.driver.switch_to.window(handle)
                        self._seen_windows.add(handle)
                        self._capture_current(bundle, reason="new_window")
                    except Exception:
                        continue

                state = self._page_state_signature(bundle)
                if state and state != last_state:
                    last_state = state
                    self._capture_current(bundle, reason="state_change")

                if self._manual_capture_event.is_set():
                    self._manual_capture_event.clear()
                    self._capture_current(bundle, reason="manual_input")

                time.sleep(max(0.2, self.config.manual.poll_seconds))
            return 0
        except KeyboardInterrupt:
            self._stop_event.set()
            return 130

    def _candidate_clicks_for_page(self, bundle: Any, snapshot: PageSnapshot) -> list[SelectorCandidate]:
        candidates: list[SelectorCandidate] = []
        for element in snapshot.elements:
            if not element.visible or not element.enabled:
                continue
            selector = element.css_selector or element.relative_xpath or element.absolute_xpath
            if not selector:
                continue

            if _blocked_any(element.text, selector, self.config.controlled.blocked_texts, self.config.controlled.blocked_selectors):
                continue

            if element.in_form and not _allowed_selector_match(selector, self.config.controlled.allowed_click_selectors):
                continue

            if self.config.controlled.allowed_click_selectors and not _allowed_selector_match(selector, self.config.controlled.allowed_click_selectors):
                if not _allowed_text_match(element.text or element.label, self.config.controlled.allowed_menu_texts):
                    continue
            elif self.config.controlled.allowed_menu_texts and not _allowed_text_match(element.text or element.label, self.config.controlled.allowed_menu_texts):
                continue

            if element.tag not in {"a", "button", "input", "select", "label", "span", "div"}:
                continue

            for candidate in element.selector_candidates:
                if candidate.unique and candidate.score >= 50:
                    candidates.append(candidate)
                    break
            else:
                candidates.append(
                    SelectorCandidate(
                        strategy="snapshot",
                        by="css selector" if element.css_selector else "xpath",
                        selector=element.css_selector or element.relative_xpath or element.absolute_xpath,
                        base_score=50.0,
                        reason="selector derivado do snapshot",
                        score=float(element.selector_confidence),
                    )
                )
        unique: list[SelectorCandidate] = []
        seen: set[tuple[str, str]] = set()
        for item in candidates:
            key = (item.by, item.selector)
            if key in seen:
                continue
            seen.add(key)
            unique.append(item)
        return unique

    def _click_candidate(self, bundle: Any, candidate: SelectorCandidate) -> bool:
        if is_dangerous_selector(candidate.selector):
            return False
        try:
            loc = _coerce_locator({"by": candidate.by, "value": candidate.selector}, None)
            if not loc:
                return False
            if candidate.by == "css selector":
                bundle.a.click(loc)
            else:
                bundle.a.click_js(loc)
            return True
        except Exception as exc:
            log.debug("Falha a clicar candidato: %s", exc)
            return False

    def run_controlled(self, bundle: Any) -> int:
        self._login(bundle)

        visited: set[str] = set()
        depth = 0
        pages_seen = 0

        try:
            while pages_seen < self.config.controlled.max_pages and depth <= self.config.controlled.max_depth:
                snapshot = self._capture_current(bundle, reason=f"controlled_depth_{depth}")
                if snapshot is None:
                    break
                pages_seen += 1
                visited.add(snapshot.signature)

                candidates = self._candidate_clicks_for_page(bundle, snapshot)
                if not candidates:
                    break

                progressed = False
                for candidate in candidates:
                    if is_dangerous_text(candidate.selector) or is_dangerous_selector(candidate.selector):
                        continue
                    before_handles = list(bundle.driver.window_handles)
                    before_url = normalize_url(getattr(bundle.driver, "current_url", ""))
                    if not self._click_candidate(bundle, candidate):
                        continue

                    try:
                        self.inventory._wait_ready(bundle.driver, timeout_seconds=self.config.controlled.timeout_seconds, stable_seconds=self.config.controlled.stability_seconds)
                    except TimeoutException:
                        pass

                    after_snapshot = self._capture_current(bundle, reason="controlled_click")
                    after_sig = after_snapshot.signature if after_snapshot else ""
                    if after_sig and after_sig not in visited:
                        visited.add(after_sig)
                        progressed = True

                    try:
                        after_handles = list(bundle.driver.window_handles)
                    except Exception:
                        after_handles = before_handles

                    if len(after_handles) > len(before_handles):
                        new_handles = [handle for handle in after_handles if handle not in before_handles]
                        for handle in new_handles:
                            try:
                                bundle.driver.switch_to.window(handle)
                                self._capture_current(bundle, reason="new_window")
                                bundle.driver.close()
                            except Exception:
                                pass
                        try:
                            bundle.driver.switch_to.window(before_handles[0])
                        except Exception:
                            pass
                    elif normalize_url(getattr(bundle.driver, "current_url", "")) != before_url:
                        try:
                            bundle.driver.back()
                            self.inventory._wait_ready(bundle.driver, timeout_seconds=10, stable_seconds=0.5)
                        except Exception:
                            pass

                    if after_sig and len(visited) >= self.config.controlled.max_pages:
                        break

                if not progressed:
                    break
                depth += 1

            return 0
        except KeyboardInterrupt:
            self._stop_event.set()
            return 130

    def run_record(self, bundle: Any) -> int:
        self._login(bundle)

        try:
            process_name = input("Nome do processo a gravar: ").strip()
        except EOFError:
            process_name = ""
        process_name = process_name or f"record_{self.run_id}"

        log_kv(log, "[record] A instalar recorder", level=logging.INFO, run_id=self.run_id, process=process_name)
        session = InteractionRecorder(
            process_name=process_name,
            root=self.root,
            site_recorder=self.recorder,
            dom_inventory=self.inventory,
            capture_timeout_seconds=self.config.record.initial_capture_timeout_seconds,
            poll_seconds=self.config.record.poll_seconds,
            capture_hidden=self.config.record.capture_hidden,
        )
        self.record_session = session
        self._bootstrap_record_mode(bundle, session, process_name)

        previous_state = collect_page_state(bundle.driver)
        previous_signature = page_state_signature(previous_state)
        last_window_handles = set()
        try:
            last_window_handles = set(bundle.driver.window_handles)
        except Exception:
            last_window_handles = set()
        last_handle = str(getattr(bundle.driver, "current_window_handle", "") or "")

        def _append_synthetic(action: str, before_state: Mapping[str, Any], after_state: Mapping[str, Any], **extra: Any) -> None:
            payload = {
                "action": action,
                "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "page_url": after_state.get("url", before_state.get("url", "")),
                "page_title": after_state.get("title", before_state.get("title", "")),
                "window_index": 0,
                "window_handle": after_state.get("window_handle", before_state.get("window_handle", "")),
                "iframe_path": after_state.get("iframe_path", before_state.get("iframe_path", "top")),
                "before_signature": page_state_signature(before_state),
                "after_signature": page_state_signature(after_state),
                "wait_condition": "synthetic",
            }
            payload.update(extra)
            session.raw_events.append(payload)

        try:
            while not self._stop_event.is_set() and not session.stopped:
                self._drain_commands(bundle, session)
                if self._stop_event.is_set() or session.stopped:
                    break

                before_state = previous_state
                before_signature = previous_signature
                before_handle = last_handle
                try:
                    handles = list(bundle.driver.window_handles)
                except Exception:
                    handles = []
                current_handle = str(getattr(bundle.driver, "current_window_handle", "") or "")
                current_state = collect_page_state(bundle.driver)
                current_signature = page_state_signature(current_state)

                if handles and len(handles) > len(last_window_handles):
                    new_handles = [handle for handle in handles if handle not in last_window_handles]
                    for handle in new_handles:
                        _append_synthetic("new_window", before_state, current_state, new_window_handle=handle)
                    session.capture_checkpoint(bundle.driver, "new_window")

                if current_handle and current_handle != before_handle:
                    _append_synthetic("window_changed", before_state, current_state, previous_window_handle=before_handle)

                if current_state.get("url") != before_state.get("url"):
                    _append_synthetic("url_changed", before_state, current_state, previous_url=before_state.get("url", ""))

                if current_signature != before_signature:
                    _append_synthetic("dom_changed", before_state, current_state)
                    previous_signature = current_signature
                    self._capture_current(bundle, reason="record_state_change")

                if current_state.get("modal_count") != before_state.get("modal_count"):
                    action = "modal_visible" if int(current_state.get("modal_count", 0) or 0) > int(before_state.get("modal_count", 0) or 0) else "modal_hidden"
                    _append_synthetic(action, before_state, current_state, modal_count=current_state.get("modal_count", 0))

                events = session.poll(bundle.driver)
                if events:
                    log_kv(log, "Eventos gravados.", level=logging.INFO, run_id=self.run_id, count=len(events), process=process_name)

                last_window_handles = set(handles)
                last_handle = current_handle or last_handle
                previous_state = current_state
                time.sleep(max(0.15, self.config.record.poll_seconds))

            return 0
        except KeyboardInterrupt:
            self._stop_event.set()
            return 130
        finally:
            try:
                session.finalize(bundle.driver)
            except Exception:
                pass

    def finalize(self) -> dict[str, Any]:
        return self.recorder.finalize(run_id=self.run_id, settings=self.settings)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Mapeador autenticado da estrutura do site SOMA.")
    parser.add_argument("--mode", choices=("manual", "controlled", "record"), default="manual", help="Modo de execução.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help="Caminho para o site_mapper.json.")
    parser.add_argument("--run-id", default="", help="Identificador do run. Se omitido, é gerado automaticamente.")
    parser.add_argument("--headless", dest="headless", action="store_true", help="Executar o browser em headless.")
    parser.add_argument("--no-headless", dest="headless", action="store_false", help="Executar o browser com janela visível.")
    parser.set_defaults(headless=None)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    if args.mode == "record" and args.headless is True:
        parser.error("--mode record exige browser visível; remova --headless ou use --no-headless.")
    settings = _load_settings()
    config = load_site_mapper_config(args.config)
    run_id = args.run_id.strip() or new_run_id(12)

    configure_logging(settings)
    ensure_artifacts_dirs(settings)

    runner = SiteMapRunner(
        settings,
        config=config,
        run_id=run_id,
        mode=args.mode,
        headless=False if args.mode == "record" else args.headless,
    )
    bundle = None
    exit_code = 1

    with step(log, "site_mapper.start", mode=args.mode, run_id=run_id):
        try:
            bundle = runner._build_bundle()
            if args.mode == "manual":
                exit_code = runner.run_manual(bundle)
            elif args.mode == "controlled":
                exit_code = runner.run_controlled(bundle)
            else:
                exit_code = runner.run_record(bundle)
            summary = runner.finalize()
            log_kv(
                log,
                "Mapeamento concluído.",
                level=logging.INFO,
                run_id=run_id,
                mode=args.mode,
                pages=summary.get("captured_pages"),
                elements=summary.get("captured_elements"),
                root=runner.root,
            )
            return exit_code
        finally:
            try:
                if bundle is not None:
                    bundle.quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
