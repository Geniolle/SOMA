from __future__ import annotations

import hashlib
import json
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from selenium.webdriver.common.by import By

from soma_app.automation.dom_inventory import (
    DomInventory,
    PageSnapshot,
    css_selector_for_element,
    normalize_text,
    normalize_url,
    sanitize_json_value,
    sanitize_text,
    selector_candidates_for_element,
    validate_selector,
    xpath_relative_for_element,
)
from soma_app.infra.trace import log_kv

log = logging.getLogger("soma_app.automation.interaction_recorder")

_ALLOWED_DATA_ATTRS = {
    "data-testid",
    "data-test",
    "data-qa",
    "data-cy",
    "data-id",
    "data-name",
}
_TEXT_TAGS = {"a", "button", "label", "option", "span", "th"}
_STATE_ACTIONS = {
    "dom_changed",
    "url_changed",
    "modal_visible",
    "modal_hidden",
    "window_changed",
    "new_window",
    "window_closed",
}
_SENSITIVE_QUERY_KEYS = re.compile(
    r"(?i)(token|auth|authorization|password|passwd|secret|session|sid|key|code|"
    r"email|user|username|document|nif|vat|iban|account)"
)


DEFAULT_INTERACTION_SCRIPT = r"""
(function () {
  if (window.__SOMA_INTERACTION_RECORDER__ && window.__SOMA_INTERACTION_RECORDER__.installed) {
    return true;
  }

  const recorder = {
    installed: true,
    seq: 0,
    queue: [],
    lastModalVisible: false,
    lastUrl: location.href,
  };

  const safeText = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

  const labelFor = (el) => {
    try {
      if (el.labels && el.labels.length) {
        return Array.from(el.labels)
          .map((label) => safeText(label.innerText || label.textContent || ""))
          .filter(Boolean)
          .join(" | ");
      }
    } catch (err) {}
    const id = el && el.getAttribute ? el.getAttribute("id") : "";
    if (!id) return "";
    try {
      const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
      return label ? safeText(label.innerText || label.textContent || "") : "";
    } catch (err) {
      return "";
    }
  };

  const stableDataAttrs = (el) => {
    const out = {};
    const allowed = new Set(["data-testid", "data-test", "data-qa", "data-cy", "data-id", "data-name"]);
    try {
      Array.from(el.attributes || []).forEach((attr) => {
        if (allowed.has(attr.name)) out[attr.name] = String(attr.value || "");
      });
    } catch (err) {}
    return out;
  };

  const describe = (el) => {
    if (!el || !el.getAttribute) return {};
    const rect = el.getBoundingClientRect ? el.getBoundingClientRect() : { x: 0, y: 0, width: 0, height: 0 };
    const style = window.getComputedStyle ? window.getComputedStyle(el) : null;
    const tag = (el.tagName || "").toLowerCase();
    const allowedTextTags = new Set(["a", "button", "label", "option", "span", "th"]);
    return {
      tag,
      id: el.getAttribute("id") || "",
      name: el.getAttribute("name") || "",
      type: el.getAttribute("type") || "",
      class_name: el.getAttribute("class") || "",
      role: el.getAttribute("role") || "",
      label: labelFor(el),
      aria_label: el.getAttribute("aria-label") || "",
      aria_labelledby: el.getAttribute("aria-labelledby") || "",
      placeholder: el.getAttribute("placeholder") || "",
      title: el.getAttribute("title") || "",
      text: allowedTextTags.has(tag) ? safeText(el.innerText || el.textContent || "").slice(0, 120) : "",
      data_attrs: stableDataAttrs(el),
      visible: !!((rect.width || rect.height) && (!style || (
        style.display !== "none" && style.visibility !== "hidden" && style.opacity !== "0"
      ))),
      enabled: !el.disabled,
      x: Math.round(rect.x || 0),
      y: Math.round(rect.y || 0),
      width: Math.round(rect.width || 0),
      height: Math.round(rect.height || 0),
      in_form: !!(el.closest && el.closest("form")),
    };
  };

  const modalCount = () => {
    try {
      return document.querySelectorAll(
        '[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container'
      ).length;
    } catch (err) {
      return 0;
    }
  };

  const push = (action, el, extra) => {
    recorder.seq += 1;
    recorder.queue.push({
      seq: recorder.seq,
      timestamp: new Date().toISOString(),
      action,
      page_url: location.href,
      page_title: document.title || "",
      ...describe(el),
      ...(extra || {}),
    });
  };

  const scanModalTransitions = () => {
    const visible = modalCount() > 0;
    if (visible && !recorder.lastModalVisible) {
      recorder.lastModalVisible = true;
      push("modal_visible", document.activeElement || document.body, { modal_count: modalCount() });
    } else if (!visible && recorder.lastModalVisible) {
      recorder.lastModalVisible = false;
      push("modal_hidden", document.activeElement || document.body, { modal_count: 0 });
    }
  };

  const isSelect2 = (el) => {
    try {
      const cls = String((el && el.className) || "");
      return cls.includes("select2") || !!(el && el.closest && el.closest(".select2-container"));
    } catch (err) {
      return false;
    }
  };

  document.addEventListener("click", (event) => {
    const target = event.target || document.activeElement || document.body;
    push("click", target, { is_select2: isSelect2(target) });
  }, true);

  document.addEventListener("input", (event) => {
    const target = event.target || document.activeElement || document.body;
    const tag = (target.tagName || "").toLowerCase();
    const type = String(target.type || "").toLowerCase();
    const select2 = isSelect2(target) || String(target.className || "").includes("select2-search__field");
    const extra = {
      value: "[redacted]",
      value_length: String(target.value || "").length,
      value_type: type || tag || "text",
      is_select2: select2,
    };
    if (tag === "input" && (type === "checkbox" || type === "radio")) {
      delete extra.value;
      delete extra.value_length;
      delete extra.value_type;
      extra.checked = !!target.checked;
    }
    push(select2 ? "select2_input" : "input", target, extra);
  }, true);

  document.addEventListener("change", (event) => {
    const target = event.target || document.activeElement || document.body;
    const tag = (target.tagName || "").toLowerCase();
    const type = String(target.type || "").toLowerCase();
    const select2 = isSelect2(target);
    const extra = {
      value: "[redacted]",
      value_length: String(target.value || "").length,
      value_type: type || tag || "text",
      is_select2: select2,
    };
    if (tag === "select") {
      extra.selected_index = Number.isInteger(target.selectedIndex) ? target.selectedIndex : null;
      extra.option_count = target.options ? target.options.length : null;
    }
    if (tag === "input" && (type === "checkbox" || type === "radio")) {
      delete extra.value;
      delete extra.value_length;
      delete extra.value_type;
      extra.checked = !!target.checked;
    }
    push(select2 ? "select2_change" : (tag === "select" ? "select" : "change"), target, extra);
  }, true);

  document.addEventListener("submit", (event) => {
    push("submit", event.target || document.activeElement || document.body, {});
  }, true);

  document.addEventListener("keydown", (event) => {
    if (String(event.key || "").toLowerCase() === "enter") {
      push("enter", event.target || document.activeElement || document.body, {});
    }
  }, true);

  const originalOpen = window.open;
  try {
    window.open = function () {
      push("new_window_requested", document.activeElement || document.body, {
        requested_url: String(arguments[0] || ""),
      });
      return originalOpen.apply(window, arguments);
    };
  } catch (err) {}

  const wrapDialog = (name) => {
    const original = window[name];
    if (typeof original !== "function") return;
    try {
      window[name] = function () {
        const message = String(arguments[0] || "");
        push(name, document.activeElement || document.body, {
          message: "[redacted]",
          message_length: message.length,
        });
        return original.apply(window, arguments);
      };
    } catch (err) {}
  };
  wrapDialog("alert");
  wrapDialog("confirm");
  wrapDialog("prompt");

  const observer = new MutationObserver(() => scanModalTransitions());
  observer.observe(document.documentElement || document.body, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeFilter: ["class", "style", "aria-hidden", "aria-modal", "open"],
  });

  recorder.flush = function () {
    if (location.href !== recorder.lastUrl) {
      recorder.lastUrl = location.href;
      push("url_changed", document.activeElement || document.body, { url: location.href });
    }
    scanModalTransitions();
    const out = recorder.queue.slice();
    recorder.queue.length = 0;
    return out;
  };

  recorder.discard = function () {
    recorder.queue.length = 0;
    recorder.lastUrl = location.href;
    scanModalTransitions();
    recorder.queue.length = 0;
    return true;
  };

  scanModalTransitions();
  window.__SOMA_INTERACTION_RECORDER__ = recorder;
  return true;
})();
"""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _hash_text(value: str) -> str:
    return hashlib.sha256((value or "").encode("utf-8")).hexdigest()


def _safe_filename(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in {"-", "_", "."} else "_" for ch in str(value or ""))
    return cleaned.strip("._") or "record"


def _sanitize_url(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        parts = urlsplit(raw)
        query = []
        for key, val in parse_qsl(parts.query, keep_blank_values=True):
            query.append((key, "[redacted]" if _SENSITIVE_QUERY_KEYS.search(key) else val))
        return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), ""))
    except Exception:
        return sanitize_text(raw)


def collect_page_state(driver: Any) -> dict[str, Any]:
    script = """
    const main = document.querySelector("main, [role='main'], #content, .content, .main")
      || document.body || document.documentElement;
    const text = main ? (main.innerText || main.textContent || "") : "";
    const html = main ? (main.innerHTML || "") : "";
    return {
      url: location.href,
      title: document.title || "",
      interactive_count: document.querySelectorAll(
        "input, select, textarea, button, a, [role], [contenteditable]"
      ).length,
      modal_count: document.querySelectorAll(
        '[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container'
      ).length,
      frame_count: document.querySelectorAll("iframe").length,
      text_hash: String(Array.from(text).reduce((a, c) => ((a << 5) - a + c.charCodeAt(0)) | 0, 0) >>> 0),
      html_hash: String(Array.from(html).reduce((a, c) => ((a << 5) - a + c.charCodeAt(0)) | 0, 0) >>> 0),
      alerts_count: 0
    };
    """
    state = driver.execute_script(script) or {}
    try:
        handles = list(driver.window_handles)
        handle = str(driver.current_window_handle)
        state["window_handle"] = handle
        state["window_index"] = handles.index(handle) if handle in handles else 0
    except Exception:
        state.setdefault("window_handle", "")
        state.setdefault("window_index", 0)
    state.setdefault("iframe_path", "top")
    state["url"] = _sanitize_url(state.get("url", ""))
    state["title"] = sanitize_text(state.get("title", ""))
    return state


def page_state_signature(state: Mapping[str, Any]) -> str:
    payload = "|".join(
        [
            normalize_url(str(state.get("url", ""))),
            sanitize_text(state.get("title", "")),
            str(state.get("window_handle", "")),
            str(state.get("window_index", 0)),
            str(state.get("iframe_path", "")),
            str(state.get("interactive_count", 0)),
            str(state.get("modal_count", 0)),
            str(state.get("frame_count", 0)),
            str(state.get("text_hash", "")),
            str(state.get("html_hash", "")),
            str(state.get("alerts_count", 0)),
        ]
    )
    return _hash_text(payload)


def sanitize_record_value(
    action: str,
    value: Any = None,
    *,
    value_length: int | None = None,
    value_type: str = "text",
    checked: bool | None = None,
    selected_index: int | None = None,
    option_count: int | None = None,
) -> dict[str, Any]:
    action_norm = normalize_text(action)
    if checked is not None or (
        action_norm in {"click", "change", "input"} and value_type in {"checkbox", "radio"}
    ):
        return {"checked": bool(checked)}

    payload: dict[str, Any] = {}
    if action_norm in {"select", "select2_choose", "input", "change", "select2_input", "select2_change"}:
        payload["value"] = "[redacted]"
        if value_length is not None:
            payload["value_length"] = max(0, int(value_length))
        if value_type:
            payload["value_type"] = str(value_type)
        if selected_index is not None:
            payload["selected_index"] = int(selected_index)
        if option_count is not None:
            payload["option_count"] = max(0, int(option_count))
    elif value is not None:
        payload["value"] = "[redacted]"
    return payload


def _event_key(event: Mapping[str, Any]) -> tuple[str, str, str, str, str]:
    return (
        normalize_text(event.get("action")),
        normalize_text(event.get("page_url")),
        normalize_text(event.get("window_handle")),
        normalize_text(event.get("iframe_path")),
        normalize_text(
            event.get("id")
            or event.get("name")
            or event.get("label")
            or event.get("aria_label")
            or event.get("css_selector")
            or event.get("relative_xpath")
            or event.get("absolute_xpath")
        ),
    )


def _event_timestamp(event: Mapping[str, Any], fallback: float = 0.0) -> float:
    try:
        return datetime.fromisoformat(str(event.get("timestamp") or "").replace("Z", "+00:00")).timestamp()
    except Exception:
        return fallback


def dedupe_input_events(
    events: Sequence[Mapping[str, Any]],
    *,
    max_gap_seconds: float = 0.8,
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    last_index: dict[tuple[str, str, str, str, str], int] = {}
    last_ts: dict[tuple[str, str, str, str, str], float] = {}
    for event in events:
        item = dict(event)
        if normalize_text(item.get("action")) not in {"input", "select2_input"}:
            out.append(item)
            continue
        key = _event_key(item)
        ts = _event_timestamp(item, float(len(out)))
        prev_idx = last_index.get(key)
        prev_ts = last_ts.get(key)
        if prev_idx is not None and prev_ts is not None and ts - prev_ts <= max_gap_seconds:
            out[prev_idx] = item
        else:
            last_index[key] = len(out)
            out.append(item)
        last_ts[key] = ts
    return out


def consolidate_select2_events(events: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    buffer: list[dict[str, Any]] = []
    buffer_key: tuple[str, str, str] | None = None

    def flush() -> None:
        nonlocal buffer, buffer_key
        if not buffer:
            return
        first, last = buffer[0], buffer[-1]
        merged = dict(last)
        merged["action"] = "select2_choose"
        merged["field"] = first.get("field") or first.get("label") or first.get("name") or first.get("id") or ""
        merged["wait_condition"] = merged.get("wait_condition") or "element_value_changed|modal_hidden|dom_changed"
        merged["selected_value"] = "[redacted]"
        merged["opener_selector"] = first.get("css_selector") or first.get("relative_xpath") or first.get("absolute_xpath") or ""
        merged["search_selector"] = next(
            (
                item.get("css_selector") or item.get("relative_xpath") or ""
                for item in buffer
                if normalize_text(item.get("action")) in {"input", "select2_input"}
            ),
            "",
        )
        merged["result_selector"] = last.get("css_selector") or last.get("relative_xpath") or last.get("absolute_xpath") or ""
        out.append(merged)
        buffer = []
        buffer_key = None

    for event in events:
        item = dict(event)
        if bool(item.get("is_select2")) or "select2" in normalize_text(item.get("action")):
            key = (
                normalize_text(item.get("page_url")),
                normalize_text(item.get("window_handle")),
                normalize_text(item.get("iframe_path")),
            )
            if buffer_key is None or key == buffer_key:
                buffer_key = key
                buffer.append(item)
                continue
            flush()
            buffer_key = key
            buffer.append(item)
            continue
        flush()
        out.append(item)
    flush()
    return out


def dedupe_state_events(
    events: Sequence[Mapping[str, Any]],
    *,
    max_gap_seconds: float = 1.0,
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    last_by_key: dict[tuple[str, str, str, str, str], tuple[int, float]] = {}
    for event in events:
        item = dict(event)
        action = normalize_text(item.get("action"))
        if action not in _STATE_ACTIONS:
            out.append(item)
            continue
        key = (
            action,
            normalize_text(item.get("window_handle")),
            normalize_text(item.get("iframe_path")),
            normalize_text(item.get("before_signature")),
            normalize_text(item.get("after_signature")),
        )
        ts = _event_timestamp(item, float(len(out)))
        previous = last_by_key.get(key)
        if previous and ts - previous[1] <= max_gap_seconds:
            out[previous[0]] = item
            last_by_key[key] = (previous[0], ts)
            continue
        last_by_key[key] = (len(out), ts)
        out.append(item)
    return out


def consolidate_events(events: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    consolidated = consolidate_select2_events(dedupe_input_events(events))
    consolidated = dedupe_state_events(consolidated)
    final: list[dict[str, Any]] = []
    previous: dict[str, Any] | None = None
    for event in consolidated:
        item = dict(event)
        action = normalize_text(item.get("action"))
        if previous and action == normalize_text(previous.get("action")) and _event_key(item) == _event_key(previous):
            if action in {"input", "change", "select"}:
                previous.update(item)
                continue
        final.append(item)
        previous = item
    return final


def suggest_wait_condition(before: Mapping[str, Any], after: Mapping[str, Any]) -> str:
    conditions: list[str] = []
    if normalize_url(str(before.get("url", ""))) != normalize_url(str(after.get("url", ""))):
        conditions.append("url_changed")
    if (
        str(before.get("window_handle", "")) != str(after.get("window_handle", ""))
        or int(before.get("window_index", 0) or 0) != int(after.get("window_index", 0) or 0)
    ):
        conditions.append("window_changed")
    if str(before.get("iframe_path", "")) != str(after.get("iframe_path", "")):
        conditions.append("iframe_changed")
    if int(before.get("modal_count", 0) or 0) < int(after.get("modal_count", 0) or 0):
        conditions.append("modal_visible")
    if int(before.get("modal_count", 0) or 0) > int(after.get("modal_count", 0) or 0):
        conditions.append("modal_hidden")
    if int(before.get("interactive_count", 0) or 0) != int(after.get("interactive_count", 0) or 0):
        conditions.append("element_count_changed")
    if (
        str(before.get("html_hash", "")) != str(after.get("html_hash", ""))
        or str(before.get("text_hash", "")) != str(after.get("text_hash", ""))
    ):
        conditions.append("dom_changed")
    if int(before.get("value_length", -1) or -1) != int(after.get("value_length", -1) or -1):
        conditions.append("element_value_changed")
    return "|".join(conditions) if conditions else "dom_changed"


def build_selector_candidate_payload(
    driver: Any,
    event: Mapping[str, Any],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    element = {
        "tag": sanitize_text(event.get("tag")),
        "type": sanitize_text(event.get("type")),
        "id": sanitize_text(event.get("id")),
        "name": sanitize_text(event.get("name")),
        "class_name": sanitize_text(event.get("class_name")),
        "role": sanitize_text(event.get("role")),
        "label": sanitize_text(event.get("label")),
        "aria_label": sanitize_text(event.get("aria_label")),
        "placeholder": sanitize_text(event.get("placeholder")),
        "text": sanitize_text(event.get("text")),
        "data_attrs": {
            str(k): sanitize_text(v)
            for k, v in dict(event.get("data_attrs") or {}).items()
            if str(k) in _ALLOWED_DATA_ATTRS
        },
    }
    candidates = selector_candidates_for_element(element, label_text=element.get("label", ""))
    payload: list[dict[str, Any]] = []
    recommended: dict[str, Any] | None = None
    for candidate in candidates:
        if driver is not None:
            validation = validate_selector(driver, candidate)
            item = {
                "strategy": candidate.strategy,
                "by": candidate.by,
                "selector": candidate.selector,
                "count": validation.count,
                "unique": validation.unique,
                "score": validation.score,
                "reason": validation.reason,
            }
        else:
            item = {
                "strategy": candidate.strategy,
                "by": candidate.by,
                "selector": candidate.selector,
                "count": 0,
                "unique": False,
                "score": candidate.score,
                "reason": candidate.reason,
            }
        payload.append(item)
        if recommended is None:
            recommended = item
        elif bool(item["unique"]) and not bool(recommended["unique"]):
            recommended = item
        elif bool(item["unique"]) == bool(recommended["unique"]) and float(item["score"]) > float(recommended["score"]):
            recommended = item
    return payload, recommended or {}


def event_to_step(
    event: Mapping[str, Any],
    *,
    step_number: int,
    before_state: Mapping[str, Any],
    after_state: Mapping[str, Any],
) -> dict[str, Any]:
    element = {
        "tag": sanitize_text(event.get("tag")),
        "type": sanitize_text(event.get("type")),
        "id": sanitize_text(event.get("id")),
        "name": sanitize_text(event.get("name")),
        "class_name": sanitize_text(event.get("class_name")),
        "role": sanitize_text(event.get("role")),
        "label": sanitize_text(event.get("label")),
        "aria_label": sanitize_text(event.get("aria_label")),
        "placeholder": sanitize_text(event.get("placeholder")),
        "text": sanitize_text(event.get("text")),
        "data_attrs": sanitize_json_value(event.get("data_attrs") or {}),
    }
    css_selector = css_selector_for_element(element)
    relative_xpath = xpath_relative_for_element(element, label_text=element.get("label", ""))
    candidates = sanitize_json_value(event.get("selector_candidates") or [])
    recommended = sanitize_json_value(event.get("selector_recommended") or {})
    if not candidates:
        generated = selector_candidates_for_element(element, label_text=element.get("label", ""))
        candidates = [
            {
                "strategy": item.strategy,
                "by": item.by,
                "selector": item.selector,
                "count": 0,
                "unique": False,
                "score": item.score,
                "reason": item.reason,
            }
            for item in generated
        ]
        recommended = candidates[0] if candidates else {}

    value_payload: dict[str, Any] = {}
    if "checked" in event:
        value_payload = {"checked": bool(event.get("checked"))}
    elif normalize_text(event.get("action")) in {
        "input",
        "change",
        "select",
        "select2_choose",
        "select2_input",
        "select2_change",
    }:
        value_payload = sanitize_record_value(
            str(event.get("action")),
            value=event.get("value"),
            value_length=int(event.get("value_length") or 0) if event.get("value_length") is not None else None,
            value_type=str(event.get("value_type") or "text"),
            selected_index=int(event.get("selected_index")) if event.get("selected_index") is not None else None,
            option_count=int(event.get("option_count")) if event.get("option_count") is not None else None,
        )

    step = {
        "step_number": step_number,
        "timestamp": sanitize_text(event.get("timestamp")),
        "action": sanitize_text(event.get("action")),
        "page_url": _sanitize_url(event.get("page_url")),
        "page_title": sanitize_text(event.get("page_title")),
        "window_index": int(event.get("window_index") or 0),
        "window_handle": sanitize_text(event.get("window_handle")),
        "iframe_path": sanitize_text(event.get("iframe_path") or "top"),
        "tag": element["tag"],
        "type": element["type"],
        "id": element["id"],
        "name": element["name"],
        "class_name": element["class_name"],
        "role": element["role"],
        "label": element["label"],
        "aria_label": element["aria_label"],
        "placeholder": element["placeholder"],
        "text": element["text"],
        "css_selector": css_selector,
        "relative_xpath": relative_xpath,
        "absolute_xpath": sanitize_text(event.get("absolute_xpath")) or relative_xpath,
        "selector_candidates": candidates,
        "selector_recommended": recommended,
        "before_signature": sanitize_text(event.get("before_signature") or page_state_signature(before_state)),
        "after_signature": sanitize_text(event.get("after_signature") or page_state_signature(after_state)),
        "wait_condition": str(event.get("wait_condition") or "").strip()
        or suggest_wait_condition(before_state, after_state),
    }
    step.update(value_payload)
    if normalize_text(step["action"]) == "select2_choose":
        step["field"] = sanitize_text(event.get("field") or element["label"] or element["name"] or element["id"])
        step["opener_selector"] = sanitize_text(event.get("opener_selector") or css_selector)
        step["search_selector"] = sanitize_text(event.get("search_selector") or "")
        step["result_selector"] = sanitize_text(event.get("result_selector") or css_selector)
    return sanitize_json_value(step)


def build_workflow_summary(process_name: str, steps: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    pages = sorted({_sanitize_url(step.get("page_url")) for step in steps if step.get("page_url")})
    windows = sorted(
        {
            f"{int(step.get('window_index') or 0)}:{step.get('window_handle')}"
            for step in steps
            if step.get("window_handle")
        }
    )
    iframes = sorted({str(step.get("iframe_path") or "top") for step in steps})
    fields = sorted(
        {
            normalize_text(step.get("field") or step.get("label") or step.get("name"))
            for step in steps
            if step.get("field") or step.get("label") or step.get("name")
        }
    )
    buttons = sorted(
        {
            normalize_text(
                step.get("label")
                or step.get("text")
                or step.get("css_selector")
                or step.get("relative_xpath")
            )
            for step in steps
            if normalize_text(step.get("action")) in {"click", "submit", "enter"}
        }
    )
    selectors = sorted(
        {
            str(selector)
            for step in steps
            for selector in (
                step.get("css_selector"),
                step.get("relative_xpath"),
                step.get("absolute_xpath"),
            )
            if selector
        }
    )
    fragile = sum(
        1
        for step in steps
        for candidate in (step.get("selector_candidates") or [])
        if not bool(candidate.get("unique")) or float(candidate.get("score") or 0.0) < 60
    )
    waits = sorted(
        {
            part
            for step in steps
            for part in str(step.get("wait_condition") or "").split("|")
            if part
        }
    )
    return {
        "process_name": process_name,
        "generated_at": utc_now_iso(),
        "step_count": len(steps),
        "pages_visited": pages,
        "windows_used": windows,
        "iframes_used": iframes,
        "fields_filled": fields,
        "buttons_clicked": buttons,
        "unique_selectors": selectors,
        "fragile_selectors": fragile,
        "wait_conditions_suggested": waits,
    }


def _sanitize_raw_event(event: Mapping[str, Any]) -> dict[str, Any]:
    item = dict(event)
    action = normalize_text(item.get("action"))
    item.pop("form_text", None)
    item.pop("outer_html", None)
    item["page_url"] = _sanitize_url(item.get("page_url"))
    if "requested_url" in item:
        item["requested_url"] = _sanitize_url(item.get("requested_url"))
    if action in {"alert", "confirm", "prompt"}:
        message = str(item.get("message") or "")
        item["message"] = "[redacted]"
        item["message_length"] = int(item.get("message_length") or len(message))
        item.pop("default_value", None)
    tag = normalize_text(item.get("tag"))
    if tag not in _TEXT_TAGS:
        item["text"] = ""
    else:
        item["text"] = sanitize_text(item.get("text"), max_len=120)
    item["data_attrs"] = {
        str(k): sanitize_text(v)
        for k, v in dict(item.get("data_attrs") or {}).items()
        if str(k) in _ALLOWED_DATA_ATTRS
    }
    if action in {"input", "change", "select", "select2_input", "select2_change", "select2_choose"}:
        item.update(
            sanitize_record_value(
                action,
                value=item.get("value"),
                value_length=int(item.get("value_length") or 0) if item.get("value_length") is not None else None,
                value_type=str(item.get("value_type") or "text"),
                checked=bool(item.get("checked")) if item.get("checked") is not None else None,
                selected_index=int(item.get("selected_index")) if item.get("selected_index") is not None else None,
                option_count=int(item.get("option_count")) if item.get("option_count") is not None else None,
            )
        )
    else:
        item.pop("value", None)
    return sanitize_json_value(item)


@dataclass
class InteractionRecorder:
    process_name: str
    root: Path
    site_recorder: Any
    dom_inventory: DomInventory
    capture_timeout_seconds: float = 20.0
    poll_seconds: float = 0.5
    capture_hidden: bool = False
    paused: bool = False
    stopped: bool = False
    raw_events: list[dict[str, Any]] = field(default_factory=list)
    page_snapshots: list[dict[str, Any]] = field(default_factory=list)
    _last_state_by_context: dict[str, dict[str, Any]] = field(default_factory=dict)
    _seen_signatures: set[str] = field(default_factory=set)
    _last_page_signature: str = ""
    _known_window_handles: set[str] = field(default_factory=set)

    def __post_init__(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        self.screenshots_dir = self.root / "screenshots"
        self.dom_dir = self.root / "dom"
        self.screenshots_dir.mkdir(parents=True, exist_ok=True)
        self.dom_dir.mkdir(parents=True, exist_ok=True)
        self.safe_process_name = _safe_filename(self.process_name)
        self.dom_inventory.capture_hidden = bool(self.capture_hidden)

    def install(self, driver: Any) -> None:
        self._visit_all_windows(driver, self._install_current_tree)

    def _install_current_tree(self, driver: Any, frame_path: str = "top") -> None:
        try:
            installed = bool(
                driver.execute_script(
                    "return !!(window.__SOMA_INTERACTION_RECORDER__ "
                    "&& window.__SOMA_INTERACTION_RECORDER__.installed);"
                )
            )
        except Exception:
            installed = False
        if not installed:
            try:
                driver.execute_script(DEFAULT_INTERACTION_SCRIPT)
            except Exception:
                return
        try:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return
        for index, frame in enumerate(frames):
            try:
                driver.switch_to.frame(frame)
                self._install_current_tree(driver, f"{frame_path}/{index}")
            except Exception:
                continue
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass

    def _visit_all_windows(self, driver: Any, callback: Any) -> None:
        try:
            original = driver.current_window_handle
        except Exception:
            original = ""
        try:
            handles = list(driver.window_handles)
        except Exception:
            handles = []
        for handle in handles:
            try:
                driver.switch_to.window(handle)
                callback(driver)
            except Exception:
                continue
        if original and original in handles:
            try:
                driver.switch_to.window(original)
            except Exception:
                pass

    def _discard_current_tree(self, driver: Any) -> None:
        self._install_current_tree(driver)
        self._discard_current_context(driver)
        try:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return
        for frame in frames:
            try:
                driver.switch_to.frame(frame)
                self._discard_current_tree(driver)
            except Exception:
                continue
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass

    @staticmethod
    def _discard_current_context(driver: Any) -> None:
        try:
            driver.execute_script(
                """
                const rec = window.__SOMA_INTERACTION_RECORDER__;
                if (rec && rec.discard) rec.discard();
                return true;
                """
            )
        except Exception:
            pass

    def _collect_current_tree(
        self,
        driver: Any,
        *,
        window_handle: str,
        window_index: int,
        frame_path: str = "top",
        discard: bool = False,
    ) -> list[dict[str, Any]]:
        self._install_current_tree(driver, frame_path)
        events: list[dict[str, Any]] = []
        if discard:
            self._discard_current_context(driver)
        else:
            try:
                raw = driver.execute_script(
                    """
                    const rec = window.__SOMA_INTERACTION_RECORDER__;
                    return rec && rec.flush ? rec.flush() : [];
                    """
                )
            except Exception:
                raw = []
            state = collect_page_state(driver)
            state["window_handle"] = window_handle
            state["window_index"] = window_index
            state["iframe_path"] = frame_path
            context_key = f"{window_handle}|{frame_path}"
            before_state = self._last_state_by_context.get(context_key, state)
            for raw_event in raw or []:
                item = dict(raw_event)
                item["window_handle"] = window_handle
                item["window_index"] = window_index
                item["iframe_path"] = frame_path
                item["page_url"] = state.get("url", item.get("page_url", ""))
                item["page_title"] = state.get("title", item.get("page_title", ""))
                item["before_signature"] = page_state_signature(before_state)
                item["after_signature"] = page_state_signature(state)
                item["wait_condition"] = suggest_wait_condition(before_state, state)
                candidates, recommended = build_selector_candidate_payload(driver, item)
                item["selector_candidates"] = candidates
                item["selector_recommended"] = recommended
                events.append(_sanitize_raw_event(item))
                before_state = state
            self._last_state_by_context[context_key] = dict(state)

        try:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return events
        for index, frame in enumerate(frames):
            child_path = f"{frame_path}/{index}"
            try:
                driver.switch_to.frame(frame)
                events.extend(
                    self._collect_current_tree(
                        driver,
                        window_handle=window_handle,
                        window_index=window_index,
                        frame_path=child_path,
                        discard=discard,
                    )
                )
            except Exception:
                continue
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass
        return events

    def _collect_all_windows(self, driver: Any, *, discard: bool = False) -> list[dict[str, Any]]:
        events: list[dict[str, Any]] = []
        try:
            original = driver.current_window_handle
        except Exception:
            original = ""
        try:
            handles = list(driver.window_handles)
        except Exception:
            handles = []

        current_handles = set(handles)
        new_handles = current_handles - self._known_window_handles
        closed_handles = self._known_window_handles - current_handles
        now = utc_now_iso()
        for handle in sorted(new_handles):
            events.append(
                {
                    "action": "new_window",
                    "timestamp": now,
                    "window_handle": handle,
                    "window_index": handles.index(handle),
                    "iframe_path": "top",
                    "before_signature": "",
                    "after_signature": "",
                    "wait_condition": "new_window",
                }
            )
        for handle in sorted(closed_handles):
            events.append(
                {
                    "action": "window_closed",
                    "timestamp": now,
                    "window_handle": handle,
                    "window_index": 0,
                    "iframe_path": "top",
                    "before_signature": "",
                    "after_signature": "",
                    "wait_condition": "window_changed",
                }
            )
        self._known_window_handles = current_handles

        for index, handle in enumerate(handles):
            try:
                driver.switch_to.window(handle)
                events.extend(
                    self._collect_current_tree(
                        driver,
                        window_handle=handle,
                        window_index=index,
                        discard=discard,
                    )
                )
            except Exception:
                continue
        if original and original in handles:
            try:
                driver.switch_to.window(original)
            except Exception:
                pass
        return [] if discard else [_sanitize_raw_event(event) for event in events]

    def _capture_snapshot(self, driver: Any, reason: str) -> PageSnapshot | None:
        try:
            page_id = f"{self.safe_process_name}_{len(self.page_snapshots) + 1:04d}"
            screenshot_path = str(self.screenshots_dir / f"{page_id}.png")
            try:
                driver.save_screenshot(screenshot_path)
            except Exception:
                screenshot_path = ""
            snapshot = self.dom_inventory.capture(
                driver,
                page_id=page_id,
                screenshot_path=screenshot_path,
                dom_dir=str(self.dom_dir),
                reason=reason,
                wait_timeout_seconds=max(0.25, float(self.capture_timeout_seconds)),
                wait_stable_seconds=max(0.1, float(self.capture_timeout_seconds) / 20.0),
            )
            if snapshot.signature in self._seen_signatures:
                return None
            self._seen_signatures.add(snapshot.signature)
            if hasattr(self.site_recorder, "save_page"):
                self.site_recorder.save_page(snapshot)
            self.page_snapshots.append(snapshot.to_dict())
            log_kv(
                log,
                "Snapshot gravado.",
                level=logging.INFO,
                reason=reason,
                signature=snapshot.signature,
                page=snapshot.url,
            )
            return snapshot
        except Exception as exc:
            log.warning("Falha a gravar snapshot: %s", exc)
            return None

    def capture_checkpoint(self, driver: Any, label: str, *, timeout_seconds: float | None = None) -> dict[str, Any]:
        state = collect_page_state(driver)
        previous_timeout = self.capture_timeout_seconds
        if timeout_seconds is not None:
            self.capture_timeout_seconds = max(1.0, float(timeout_seconds))
        try:
            snapshot = self._capture_snapshot(driver, reason=label)
        finally:
            self.capture_timeout_seconds = previous_timeout
        payload = _sanitize_raw_event(
            {
                "action": "checkpoint",
                "label": sanitize_text(label),
                "timestamp": utc_now_iso(),
                "page_url": state.get("url", ""),
                "page_title": state.get("title", ""),
                "window_index": state.get("window_index", 0),
                "window_handle": state.get("window_handle", ""),
                "iframe_path": "top",
                "before_signature": page_state_signature(state),
                "after_signature": snapshot.signature if snapshot else page_state_signature(state),
            }
        )
        self.raw_events.append(payload)
        return payload

    def record_marker(self, driver: Any, label: str) -> dict[str, Any]:
        state = collect_page_state(driver)
        payload = _sanitize_raw_event(
            {
                "action": "marker",
                "label": sanitize_text(label),
                "timestamp": utc_now_iso(),
                "page_url": state.get("url", ""),
                "page_title": state.get("title", ""),
                "window_index": state.get("window_index", 0),
                "window_handle": state.get("window_handle", ""),
                "iframe_path": "top",
                "before_signature": page_state_signature(state),
                "after_signature": page_state_signature(state),
            }
        )
        self.raw_events.append(payload)
        return payload

    def record_checkpoint_event(self, driver: Any, label: str) -> dict[str, Any]:
        state = collect_page_state(driver)
        payload = _sanitize_raw_event(
            {
                "action": "checkpoint",
                "label": sanitize_text(label),
                "timestamp": utc_now_iso(),
                "page_url": state.get("url", ""),
                "page_title": state.get("title", ""),
                "window_index": state.get("window_index", 0),
                "window_handle": state.get("window_handle", ""),
                "iframe_path": state.get("iframe_path", "top"),
                "before_signature": page_state_signature(state),
                "after_signature": page_state_signature(state),
            }
        )
        self.raw_events.append(payload)
        return payload

    def pause(self, driver: Any | None = None) -> None:
        if driver is not None:
            try:
                current_context = getattr(driver, "current_context", None)
                if isinstance(current_context, dict) and "queue" in current_context:
                    current_context["queue"] = []
            except Exception:
                pass
        self.paused = True
        if driver is not None:
            self.record_marker(driver, "pause")

    def resume(self, driver: Any | None = None) -> None:
        if driver is not None:
            self._discard_current_context(driver)
        self.paused = False
        if driver is not None:
            self.record_marker(driver, "resume")

    def request_stop(self, driver: Any | None = None) -> None:
        if driver is not None:
            self.flush_final(driver)
            self.record_marker(driver, "stop")
        self.stopped = True

    def process_command(self, command: str, driver: Any) -> str:
        cmd = normalize_text(command)
        if not cmd:
            self.record_checkpoint_event(driver, "checkpoint")
            return "checkpoint"
        if cmd in {"q", "quit", "exit", "done", "finish"}:
            self.request_stop(driver)
            return "stop"
        if cmd == "pause":
            self.pause(driver)
            return "pause"
        if cmd == "resume":
            self.resume(driver)
            return "resume"
        if cmd == "mark":
            self.record_marker(driver, "mark")
            return "mark"
        self.record_marker(driver, cmd)
        return cmd

    def ingest(self, driver: Any, events: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
        if self.paused or self.stopped:
            return []
        accepted: list[dict[str, Any]] = []
        for event in events:
            item = _sanitize_raw_event(event)
            if not item.get("selector_candidates") and driver is not None:
                candidates, recommended = build_selector_candidate_payload(driver, item)
                item["selector_candidates"] = candidates
                item["selector_recommended"] = recommended
            accepted.append(item)
            self.raw_events.append(item)
        return accepted

    def poll(self, driver: Any) -> list[dict[str, Any]]:
        events = self._collect_all_windows(driver, discard=self.paused)
        if self.paused or self.stopped:
            return []
        accepted = self.ingest(driver, events)
        try:
            state = collect_page_state(driver)
            signature = page_state_signature(state)
            if signature and signature != self._last_page_signature:
                self._last_page_signature = signature
                self._capture_snapshot(driver, reason="dom_changed")
        except Exception:
            pass
        return accepted

    def flush_final(self, driver: Any) -> list[dict[str, Any]]:
        was_stopped = self.stopped
        was_paused = self.paused
        self.stopped = False
        self.paused = False
        try:
            events = self._collect_all_windows(driver, discard=False)
            return self.ingest(driver, events)
        finally:
            self.stopped = was_stopped
            self.paused = was_paused

    def finalize(self, driver: Any | None = None) -> dict[str, Any]:
        status_path = self.root / "record_status.json"
        self._write_json(
            "record_status.json",
            {"status": "finalizing", "process_name": self.process_name, "updated_at": utc_now_iso()},
        )
        try:
            if driver is not None and not self.stopped:
                self.flush_final(driver)
            workflow = consolidate_events(self.raw_events)
            steps: list[dict[str, Any]] = []
            previous_state: dict[str, Any] = {}
            for index, event in enumerate(workflow, start=1):
                state = {
                    "url": event.get("page_url", ""),
                    "title": event.get("page_title", ""),
                    "window_handle": event.get("window_handle", ""),
                    "window_index": event.get("window_index", 0),
                    "iframe_path": event.get("iframe_path", "top"),
                    "interactive_count": event.get("interactive_count", 0),
                    "modal_count": event.get("modal_count", 0),
                    "frame_count": event.get("frame_count", 0),
                    "text_hash": event.get("text_hash", ""),
                    "html_hash": event.get("html_hash", ""),
                    "value_length": event.get("value_length"),
                }
                steps.append(
                    event_to_step(
                        event,
                        step_number=index,
                        before_state=previous_state or state,
                        after_state=state,
                    )
                )
                previous_state = state

            summary = build_workflow_summary(self.process_name, steps)
            elements_used = self._build_elements_used(steps)
            locator_candidates = self._build_locator_candidates_used(steps)
            self._write_json("steps.json", self.raw_events)
            self._write_json("workflow.json", {"process_name": self.process_name, "steps": steps})
            self._write_json("workflow_summary.json", summary)
            self._write_json("elements_used.json", elements_used)
            self._write_json("locator_candidates_used.json", locator_candidates)
            self._write_json(
                "record_status.json",
                {
                    "status": "complete",
                    "process_name": self.process_name,
                    "updated_at": utc_now_iso(),
                    "step_count": len(steps),
                },
            )
            return {
                "steps": steps,
                "workflow_summary": summary,
                "elements_used": elements_used,
                "locator_candidates_used": locator_candidates,
            }
        except Exception as exc:
            log.exception("Falha ao finalizar a gravação: %s", exc)
            try:
                status_path.write_text(
                    json.dumps(
                        {
                            "status": "failed",
                            "process_name": self.process_name,
                            "updated_at": utc_now_iso(),
                            "error": sanitize_text(exc),
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )
            except Exception:
                pass
            raise

    def _write_json(self, name: str, payload: Any) -> None:
        path = self.root / name
        path.write_text(
            json.dumps(sanitize_json_value(payload), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _build_elements_used(self, steps: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
        used: list[dict[str, Any]] = []
        seen: set[tuple[str, str, str, str]] = set()
        for step in steps:
            key = (
                normalize_text(step.get("page_url")),
                normalize_text(step.get("iframe_path")),
                normalize_text(step.get("css_selector") or step.get("relative_xpath") or step.get("absolute_xpath")),
                normalize_text(step.get("label") or step.get("name") or step.get("id") or step.get("text")),
            )
            if key in seen:
                continue
            seen.add(key)
            used.append(
                sanitize_json_value(
                    {
                        "page_url": step.get("page_url", ""),
                        "page_title": step.get("page_title", ""),
                        "window_index": step.get("window_index", 0),
                        "window_handle": step.get("window_handle", ""),
                        "iframe_path": step.get("iframe_path", ""),
                        "tag": step.get("tag", ""),
                        "id": step.get("id", ""),
                        "name": step.get("name", ""),
                        "role": step.get("role", ""),
                        "label": step.get("label", ""),
                        "aria_label": step.get("aria_label", ""),
                        "placeholder": step.get("placeholder", ""),
                        "text": step.get("text", ""),
                        "css_selector": step.get("css_selector", ""),
                        "relative_xpath": step.get("relative_xpath", ""),
                        "absolute_xpath": step.get("absolute_xpath", ""),
                        "selector_recommended": step.get("selector_recommended", {}),
                    }
                )
            )
        return used

    def _build_locator_candidates_used(self, steps: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
        candidates: list[dict[str, Any]] = []
        seen: set[tuple[str, str, str, str]] = set()
        for step in steps:
            for candidate in step.get("selector_candidates", []) or []:
                key = (
                    normalize_text(step.get("window_handle")),
                    normalize_text(step.get("iframe_path")),
                    normalize_text(candidate.get("by")),
                    normalize_text(candidate.get("selector")),
                )
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(
                    sanitize_json_value(
                        {
                            "page_url": step.get("page_url", ""),
                            "page_title": step.get("page_title", ""),
                            "window_handle": step.get("window_handle", ""),
                            "iframe_path": step.get("iframe_path", ""),
                            "field": step.get("field", ""),
                            "strategy": candidate.get("strategy", ""),
                            "by": candidate.get("by", ""),
                            "selector": candidate.get("selector", ""),
                            "count": candidate.get("count", 0),
                            "unique": candidate.get("unique", False),
                            "score": candidate.get("score", 0),
                            "reason": candidate.get("reason", ""),
                            "recommended": candidate == step.get("selector_recommended", {}),
                        }
                    )
                )
        return candidates


__all__ = [
    "DEFAULT_INTERACTION_SCRIPT",
    "InteractionRecorder",
    "build_workflow_summary",
    "collect_page_state",
    "consolidate_events",
    "consolidate_select2_events",
    "dedupe_input_events",
    "dedupe_state_events",
    "event_to_step",
    "page_state_signature",
    "sanitize_record_value",
    "suggest_wait_condition",
]
