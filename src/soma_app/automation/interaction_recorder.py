from __future__ import annotations

import hashlib
import json
import logging
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

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

DEFAULT_INTERACTION_SCRIPT = r"""
(function () {
  if (window.__SOMA_INTERACTION_RECORDER__ && window.__SOMA_INTERACTION_RECORDER__.installed) {
    return window.__SOMA_INTERACTION_RECORDER__;
  }

  const recorder = {
    installed: true,
    seq: 0,
    queue: [],
    lastModalVisible: false,
    lastModalCount: 0,
    lastUrl: location.href,
    lastTitle: document.title || "",
    lastWindowHandle: String(window.name || ""),
  };

  const safeText = (value) => {
    const raw = String(value ?? "");
    return raw.replace(/\s+/g, " ").trim();
  };

  const escapeXpath = (value) => {
    const s = String(value ?? "");
    if (!s.includes("'")) return `'${s}'`;
    if (!s.includes('"')) return `"${s}"`;
    return "concat(" + s.split("'").map((part) => `'${part}'`).join(", \"'\", ") + ")";
  };

  const getFramePath = () => {
    const parts = [];
    let win = window;
    let depth = 0;
    while (win && win !== win.top && depth < 8) {
      const frame = win.frameElement;
      if (!frame) break;
      const id = frame.getAttribute("id") || frame.getAttribute("name");
      if (id) {
        parts.unshift(safeText(id));
      } else {
        let index = 1;
        let sib = frame.previousElementSibling;
        while (sib) {
          if (sib.tagName && sib.tagName.toLowerCase() === "iframe") index += 1;
          sib = sib.previousElementSibling;
        }
        parts.unshift(`iframe[${index}]`);
      }
      try {
        win = win.parent;
      } catch (err) {
        break;
      }
      depth += 1;
    }
    return parts.length ? `top/${parts.join("/")}` : "top";
  };

  const getLabel = (el) => {
    try {
      if (el.labels && el.labels.length) {
        return Array.from(el.labels)
          .map((label) => safeText(label.innerText || label.textContent || ""))
          .filter(Boolean)
          .join(" | ");
      }
    } catch (err) {}
    const id = el.getAttribute && el.getAttribute("id");
    if (!id) return "";
    try {
      const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
      return label ? safeText(label.innerText || label.textContent || "") : "";
    } catch (err) {
      return "";
    }
  };

  const getDataAttrs = (el) => {
    const out = {};
    try {
      Array.from(el.attributes || []).forEach((attr) => {
        if (attr.name && attr.name.startsWith("data-")) {
          out[attr.name] = String(attr.value || "");
        }
      });
    } catch (err) {}
    return out;
  };

  const describe = (el) => {
    if (!el || !el.getAttribute) {
      return {
        tag: "",
        id: "",
        name: "",
        type: "",
        class_name: "",
        role: "",
        label: "",
        aria_label: "",
        placeholder: "",
        text: "",
        data_attrs: {},
      };
    }
    const rect = el.getBoundingClientRect ? el.getBoundingClientRect() : { x: 0, y: 0, width: 0, height: 0 };
    const style = window.getComputedStyle ? window.getComputedStyle(el) : null;
    const visible = !!((rect.width || rect.height) && (!style || (style.display !== "none" && style.visibility !== "hidden" && style.opacity !== "0")));
    return {
      tag: (el.tagName || "").toLowerCase(),
      id: el.getAttribute("id") || "",
      name: el.getAttribute("name") || "",
      type: el.getAttribute("type") || "",
      class_name: el.getAttribute("class") || "",
      role: el.getAttribute("role") || "",
      label: getLabel(el),
      aria_label: el.getAttribute("aria-label") || "",
      aria_labelledby: el.getAttribute("aria-labelledby") || "",
      placeholder: el.getAttribute("placeholder") || "",
      title: el.getAttribute("title") || "",
      text: safeText(el.innerText || el.textContent || ""),
      data_attrs: getDataAttrs(el),
      visible,
      enabled: !el.disabled,
      x: Math.round(rect.x || 0),
      y: Math.round(rect.y || 0),
      width: Math.round(rect.width || 0),
      height: Math.round(rect.height || 0),
      in_form: !!(el.closest && el.closest("form")),
      form_text: el.closest && el.closest("form") ? safeText(el.closest("form").innerText || el.closest("form").textContent || "") : "",
      css_selector: "",
      relative_xpath: "",
      absolute_xpath: "",
      selector_candidates: [],
    };
  };

  const push = (action, el, extra) => {
    recorder.seq += 1;
    const desc = describe(el);
    recorder.queue.push({
      seq: recorder.seq,
      timestamp: new Date().toISOString(),
      action,
      page_url: location.href,
      page_title: document.title || "",
      window_index: 0,
      window_handle: String(window.name || ""),
      iframe_path: getFramePath(),
      ...desc,
      ...(extra || {}),
    });
  };

  const modalCount = () => {
    try {
      return document.querySelectorAll('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container').length;
    } catch (err) {
      return 0;
    }
  };

  const scanModalTransitions = () => {
    const count = modalCount();
    if (count > 0 && !recorder.lastModalVisible) {
      recorder.lastModalVisible = true;
      recorder.lastModalCount = count;
      push("modal_visible", document.activeElement || document.body, { modal_count: count });
    } else if (count === 0 && recorder.lastModalVisible) {
      recorder.lastModalVisible = false;
      recorder.lastModalCount = 0;
      push("modal_hidden", document.activeElement || document.body, { modal_count: 0 });
    } else {
      recorder.lastModalCount = count;
    }
  };

  const isSelect2 = (el) => {
    try {
      if (!el) return false;
      const cls = String(el.className || "");
      return cls.includes("select2") || !!(el.closest && el.closest(".select2-container"));
    } catch (err) {
      return false;
    }
  };

  const isSelectControl = (el) => (el && el.tagName && el.tagName.toLowerCase() === "select");

  const onClick = (event) => {
    const target = event.target || document.activeElement || document.body;
    push("click", target, {
      is_select2: isSelect2(target),
      is_select: isSelectControl(target),
    });
  };

  const onInput = (event) => {
    const target = event.target || document.activeElement || document.body;
    const tag = (target.tagName || "").toLowerCase();
    const type = String(target.type || "").toLowerCase();
    const isSelect2Field = isSelect2(target) || String(target.className || "").includes("select2-search__field");
    const payload = {
      value: "[redacted]",
      value_length: String(target.value || "").length,
      value_type: type || tag || "text",
      is_select2: isSelect2Field,
      is_select: isSelectControl(target),
    };
    if (tag === "input" && (type === "checkbox" || type === "radio")) {
      payload.checked = !!target.checked;
      delete payload.value;
      delete payload.value_length;
      delete payload.value_type;
    }
    push(isSelect2Field ? "select2_input" : "input", target, payload);
  };

  const onChange = (event) => {
    const target = event.target || document.activeElement || document.body;
    const tag = (target.tagName || "").toLowerCase();
    const type = String(target.type || "").toLowerCase();
    const payload = {
      value: "[redacted]",
      value_length: String(target.value || "").length,
      value_type: type || tag || "text",
      is_select2: isSelect2(target),
      is_select: isSelectControl(target),
    };
    if (tag === "input" && (type === "checkbox" || type === "radio")) {
      payload.checked = !!target.checked;
      delete payload.value;
      delete payload.value_length;
      delete payload.value_type;
    }
    push(isSelect2(target) ? "select2_change" : (tag === "select" ? "select" : "change"), target, payload);
  };

  const onSubmit = (event) => {
    const target = event.target || document.activeElement || document.body;
    push("submit", target, {});
  };

  const onFocus = (event) => {
    const target = event.target || document.activeElement || document.body;
    push("focus", target, {});
  };

  const onBlur = (event) => {
    const target = event.target || document.activeElement || document.body;
    push("blur", target, {});
  };

  const onKeyDown = (event) => {
    const key = String(event.key || "").toLowerCase();
    if (key === "enter") {
      const target = event.target || document.activeElement || document.body;
      push("enter", target, {});
    }
  };

  const onUrlChange = () => {
    if (location.href !== recorder.lastUrl) {
      recorder.lastUrl = location.href;
      push("url_changed", document.activeElement || document.body, { url: location.href });
    }
  };

  const onWindowOpen = window.open;
  try {
    window.open = function () {
      push("new_window_requested", document.activeElement || document.body, { requested_url: arguments[0] || "" });
      return onWindowOpen.apply(window, arguments);
    };
  } catch (err) {}

  const originalAlert = window.alert;
  const originalConfirm = window.confirm;
  const originalPrompt = window.prompt;
  try {
    window.alert = function (message) {
      push("alert", document.activeElement || document.body, { message: safeText(message) });
      return originalAlert.apply(window, arguments);
    };
  } catch (err) {}
  try {
    window.confirm = function (message) {
      push("confirm", document.activeElement || document.body, { message: safeText(message) });
      return originalConfirm.apply(window, arguments);
    };
  } catch (err) {}
  try {
    window.prompt = function (message, defaultValue) {
      push("prompt", document.activeElement || document.body, { message: safeText(message), default_value: safeText(defaultValue) });
      return originalPrompt.apply(window, arguments);
    };
  } catch (err) {}

  const observer = new MutationObserver((mutations) => {
    let modalChanged = false;
    for (const mutation of mutations) {
      for (const node of mutation.addedNodes || []) {
        if (!node || !node.querySelectorAll) continue;
        const modal = node.matches && (node.matches('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container') || node.querySelector('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container'));
        if (modal) {
          push("element_visible", node, { mutation_type: "added" });
          modalChanged = true;
        }
      }
      for (const node of mutation.removedNodes || []) {
        if (!node || !node.querySelectorAll) continue;
        const modal = node.matches && (node.matches('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container') || node.querySelector('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container'));
        if (modal) {
          push("element_hidden", node, { mutation_type: "removed" });
          modalChanged = true;
        }
      }
    }
    if (modalChanged) {
      scanModalTransitions();
    }
  });

  const install = () => {
    document.addEventListener("click", onClick, true);
    document.addEventListener("input", onInput, true);
    document.addEventListener("change", onChange, true);
    document.addEventListener("submit", onSubmit, true);
    document.addEventListener("focus", onFocus, true);
    document.addEventListener("blur", onBlur, true);
    document.addEventListener("keydown", onKeyDown, true);
    observer.observe(document.documentElement || document.body, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ["class", "style", "aria-hidden", "aria-modal", "open"],
    });
    scanModalTransitions();
  };

  recorder.flush = function () {
    onUrlChange();
    scanModalTransitions();
    const out = recorder.queue.slice();
    recorder.queue.length = 0;
    return out;
  };

  recorder.snapshot = function () {
    const main = document.querySelector("main, [role='main'], #content, .content, .main") || document.body || document.documentElement;
    const text = main ? safeText(main.innerText || main.textContent || "") : "";
    const html = main ? (main.innerHTML || "") : "";
    const interactive = document.querySelectorAll("input, select, textarea, button, a, [role], [contenteditable]").length;
    return {
      url: location.href,
      title: document.title || "",
      window_handle: String(window.name || ""),
      iframe_path: getFramePath(),
      active_tag: document.activeElement && document.activeElement.tagName ? document.activeElement.tagName.toLowerCase() : "",
      active_id: document.activeElement && document.activeElement.getAttribute ? (document.activeElement.getAttribute("id") || "") : "",
      active_name: document.activeElement && document.activeElement.getAttribute ? (document.activeElement.getAttribute("name") || "") : "",
      interactive_count: interactive,
      modal_count: modalCount(),
      frame_count: document.querySelectorAll("iframe").length,
      text_length: text.length,
      html_hash: (function () {
        let hash = 0;
        const input = String(html || "");
        for (let i = 0; i < input.length; i += 1) {
          hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
        }
        return String(hash >>> 0);
      })(),
      text_hash: (function () {
        let hash = 0;
        const input = String(text || "");
        for (let i = 0; i < input.length; i += 1) {
          hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
        }
        return String(hash >>> 0);
      })(),
    };
  };

  install();
  window.__SOMA_INTERACTION_RECORDER__ = recorder;
  return recorder;
})();
"""


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _hash_text(value: str) -> str:
    return hashlib.sha256((value or "").encode("utf-8")).hexdigest()


def collect_page_state(driver: Any) -> dict[str, Any]:
    script = """
    const main = document.querySelector("main, [role='main'], #content, .content, .main") || document.body || document.documentElement;
    const text = main ? (main.innerText || main.textContent || "") : "";
    const html = main ? (main.innerHTML || "") : "";
    const interactive = Array.from(document.querySelectorAll("input, select, textarea, button, a, [role], [contenteditable]")).map((el) => {
      const rect = el.getBoundingClientRect();
      return [
        (el.tagName || "").toLowerCase(),
        el.getAttribute("id") || "",
        el.getAttribute("name") || "",
        el.getAttribute("type") || "",
        el.getAttribute("role") || "",
        el.getAttribute("aria-label") || "",
        Math.round(rect.x || 0),
        Math.round(rect.y || 0)
      ].join(":");
    }).join("|");
    const modalCount = document.querySelectorAll('[role="dialog"], [aria-modal="true"], .modal.show, .swal2-container').length;
    return {
      url: location.href,
      title: document.title || "",
      window_handle: String(window.name || ""),
      iframe_path: (function () {
        const parts = [];
        let win = window;
        let depth = 0;
        while (win && win !== win.top && depth < 8) {
          const frame = win.frameElement;
          if (!frame) break;
          const id = frame.getAttribute("id") || frame.getAttribute("name");
          if (id) {
            parts.unshift(id);
          } else {
            let index = 1;
            let sib = frame.previousElementSibling;
            while (sib) {
              if ((sib.tagName || "").toLowerCase() === "iframe") index += 1;
              sib = sib.previousElementSibling;
            }
            parts.unshift(`iframe[${index}]`);
          }
          try {
            win = win.parent;
          } catch (err) {
            break;
          }
          depth += 1;
        }
        return parts.length ? `top/${parts.join("/")}` : "top";
      })(),
      active_tag: document.activeElement && document.activeElement.tagName ? document.activeElement.tagName.toLowerCase() : "",
      active_id: document.activeElement && document.activeElement.getAttribute ? (document.activeElement.getAttribute("id") || "") : "",
      active_name: document.activeElement && document.activeElement.getAttribute ? (document.activeElement.getAttribute("name") || "") : "",
      interactive_count: document.querySelectorAll("input, select, textarea, button, a, [role], [contenteditable]").length,
      modal_count: modalCount,
      frame_count: document.querySelectorAll("iframe").length,
      text_length: (text || "").length,
      html_hash: (function () {
        let hash = 0;
        const input = String(html || "");
        for (let i = 0; i < input.length; i += 1) {
          hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
        }
        return String(hash >>> 0);
      })(),
      text_hash: (function () {
        let hash = 0;
        const input = String(text || "");
        for (let i = 0; i < input.length; i += 1) {
          hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
        }
        return String(hash >>> 0);
      })(),
      alerts_count: 0
    };
    """
    return driver.execute_script(script) or {}


def page_state_signature(state: Mapping[str, Any]) -> str:
    payload = "|".join(
        [
            normalize_url(str(state.get("url", ""))),
            sanitize_text(state.get("title", "")),
            str(state.get("window_handle", "")),
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


def sanitize_record_value(action: str, value: Any = None, *, value_length: int | None = None, value_type: str = "text", checked: bool | None = None, selected_index: int | None = None, option_count: int | None = None) -> dict[str, Any]:
    action_norm = normalize_text(action)
    payload: dict[str, Any] = {}
    if checked is not None or action_norm in {"click", "change", "input"} and value_type in {"checkbox", "radio"}:
        payload["checked"] = bool(checked)
        return payload

    if action_norm in {"select", "select2_choose", "input", "change"}:
        payload["value"] = "[redacted]"
        if value_length is not None:
            payload["value_length"] = int(value_length)
        if value_type:
            payload["value_type"] = str(value_type)
        if selected_index is not None:
            payload["selected_index"] = int(selected_index)
        if option_count is not None:
            payload["option_count"] = int(option_count)
        return payload

    if value is not None:
        payload["value"] = "[redacted]"
    return payload


def _event_key(event: Mapping[str, Any]) -> tuple[str, str, str, str, str]:
    return (
        normalize_text(event.get("action")),
        normalize_text(event.get("page_url")),
        normalize_text(event.get("window_handle")),
        normalize_text(event.get("iframe_path")),
        normalize_text(event.get("id") or event.get("name") or event.get("label") or event.get("aria_label") or event.get("css_selector") or event.get("relative_xpath") or event.get("absolute_xpath")),
    )


def _is_input_like(event: Mapping[str, Any]) -> bool:
    return normalize_text(event.get("action")) in {"input", "change", "select", "select2_input", "select2_change", "select2_choose"}


def dedupe_input_events(events: Sequence[Mapping[str, Any]], *, max_gap_seconds: float = 0.8) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    last_input_index: dict[tuple[str, str, str, str, str], int] = {}
    last_seen_ts: dict[tuple[str, str, str, str, str], float] = {}

    for event in events:
        item = dict(event)
        action = normalize_text(item.get("action"))
        if action not in {"input", "select2_input"}:
            out.append(item)
            continue

        key = _event_key(item)
        ts_raw = str(item.get("timestamp") or "")
        try:
            ts = datetime.fromisoformat(ts_raw.replace("Z", "+00:00")).timestamp()
        except Exception:
            ts = float(len(out))

        prev_index = last_input_index.get(key)
        prev_ts = last_seen_ts.get(key)
        if prev_index is not None and prev_ts is not None and ts - prev_ts <= max_gap_seconds:
            out[prev_index] = item
        else:
            last_input_index[key] = len(out)
            out.append(item)
        last_seen_ts[key] = ts
    return out


def consolidate_select2_events(events: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    buffer: list[dict[str, Any]] = []
    buffer_key: tuple[str, str, str] | None = None

    def flush() -> None:
        nonlocal buffer, buffer_key
        if not buffer:
            return
        first = buffer[0]
        last = buffer[-1]
        merged = dict(last)
        merged["action"] = "select2_choose"
        merged["field"] = first.get("field") or first.get("label") or first.get("name") or first.get("id") or ""
        merged["wait_condition"] = merged.get("wait_condition") or "element_value_changed|modal_hidden|dom_changed"
        merged["selected_value"] = "[redacted]"
        merged["result_selector"] = last.get("css_selector") or last.get("relative_xpath") or last.get("absolute_xpath") or ""
        merged["search_selector"] = next((item.get("css_selector") or item.get("relative_xpath") or "" for item in buffer if normalize_text(item.get("action")) in {"input", "select2_input"}), "")
        merged["opener_selector"] = first.get("css_selector") or first.get("relative_xpath") or first.get("absolute_xpath") or ""
        if any(normalize_text(item.get("action")) == "change" for item in buffer):
            merged["action"] = "select2_choose"
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
            if buffer_key is None:
                buffer_key = key
                buffer.append(item)
                continue
            if key == buffer_key:
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


def consolidate_events(events: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    deduped = dedupe_input_events(events)
    consolidated = consolidate_select2_events(deduped)
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
    if str(before.get("window_handle", "")) != str(after.get("window_handle", "")) or int(before.get("window_index", 0) or 0) != int(after.get("window_index", 0) or 0):
        conditions.append("window_changed")
    if str(before.get("iframe_path", "")) != str(after.get("iframe_path", "")):
        conditions.append("iframe_changed")
    if int(before.get("modal_count", 0) or 0) < int(after.get("modal_count", 0) or 0):
        conditions.append("modal_visible")
    if int(before.get("modal_count", 0) or 0) > int(after.get("modal_count", 0) or 0):
        conditions.append("modal_hidden")
    if int(before.get("interactive_count", 0) or 0) != int(after.get("interactive_count", 0) or 0):
        conditions.append("element_count_changed")
    if str(before.get("html_hash", "")) != str(after.get("html_hash", "")) or str(before.get("text_hash", "")) != str(after.get("text_hash", "")):
        conditions.append("dom_changed")
    if int(before.get("value_length", -1) or -1) != int(after.get("value_length", -1) or -1):
        conditions.append("element_value_changed")
    return "|".join(conditions) if conditions else "dom_changed"


def build_selector_candidate_payload(driver: Any, event: Mapping[str, Any]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
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
    candidates = selector_candidates_for_element(element, label_text=element.get("label", ""))
    payload: list[dict[str, Any]] = []
    recommended: dict[str, Any] | None = None
    for candidate in candidates:
        if driver is not None:
            validation = validate_selector(driver, candidate)
            payload_item = {
                "strategy": candidate.strategy,
                "by": candidate.by,
                "selector": candidate.selector,
                "count": validation.count,
                "unique": validation.unique,
                "score": validation.score,
                "reason": validation.reason,
            }
        else:
            payload_item = {
                "strategy": candidate.strategy,
                "by": candidate.by,
                "selector": candidate.selector,
                "count": 0,
                "unique": False,
                "score": candidate.score,
                "reason": candidate.reason,
            }
        payload.append(payload_item)
        if recommended is None:
            recommended = payload_item
        elif bool(payload_item["unique"]) and not bool(recommended["unique"]):
            recommended = payload_item
        elif bool(payload_item["unique"]) == bool(recommended["unique"]) and float(payload_item["score"]) > float(recommended["score"]):
            recommended = payload_item
    return payload, recommended or {}


def event_to_step(event: Mapping[str, Any], *, step_number: int, before_state: Mapping[str, Any], after_state: Mapping[str, Any]) -> dict[str, Any]:
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
    absolute_xpath = sanitize_text(event.get("absolute_xpath")) or relative_xpath
    candidates = sanitize_json_value(event.get("selector_candidates") or [])
    recommended = sanitize_json_value(event.get("selector_recommended") or {})
    if not candidates:
        generated_candidates = selector_candidates_for_element(element, label_text=element.get("label", ""))
        candidates = [
            {
                "strategy": candidate.strategy,
                "by": candidate.by,
                "selector": candidate.selector,
                "count": 0,
                "unique": False,
                "score": candidate.score,
                "reason": candidate.reason,
            }
            for candidate in generated_candidates
        ]
        recommended = candidates[0] if candidates else {}

    value_payload = {}
    if "checked" in event:
        value_payload = {"checked": bool(event.get("checked"))}
    elif normalize_text(event.get("action")) in {"input", "change", "select", "select2_choose"}:
        value_payload = sanitize_record_value(
            str(event.get("action")),
            value=event.get("value"),
            value_length=int(event.get("value_length") or 0) if event.get("value_length") is not None else None,
            value_type=str(event.get("value_type") or "text"),
            selected_index=int(event.get("selected_index")) if event.get("selected_index") is not None else None,
            option_count=int(event.get("option_count")) if event.get("option_count") is not None else None,
        )
        value_payload.setdefault("value", "[redacted]")

    wait_condition = str(event.get("wait_condition") or "").strip() or suggest_wait_condition(before_state, after_state)
    step = {
        "step_number": step_number,
        "timestamp": sanitize_text(event.get("timestamp")),
        "action": sanitize_text(event.get("action")),
        "page_url": sanitize_text(event.get("page_url")),
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
        "absolute_xpath": absolute_xpath,
        "selector_candidates": candidates,
        "selector_recommended": recommended,
        "before_signature": sanitize_text(event.get("before_signature") or page_state_signature(before_state)),
        "after_signature": sanitize_text(event.get("after_signature") or page_state_signature(after_state)),
        "wait_condition": wait_condition,
    }
    step.update(value_payload)
    if normalize_text(step["action"]) == "select2_choose":
        step.setdefault("field", sanitize_text(event.get("field") or element["label"] or element["name"] or element["id"]))
        step.setdefault("opener_selector", sanitize_text(event.get("opener_selector") or css_selector))
        step.setdefault("search_selector", sanitize_text(event.get("search_selector") or ""))
        step.setdefault("result_selector", sanitize_text(event.get("result_selector") or css_selector))
    return sanitize_json_value(step)


def build_workflow_summary(process_name: str, steps: Sequence[Mapping[str, Any]]) -> dict[str, Any]:
    page_urls = []
    windows = []
    iframes = []
    fields = []
    buttons = []
    selectors = set()
    fragile = 0
    wait_conditions = set()
    for step in steps:
        page_urls.append(str(step.get("page_url", "")))
        windows.append((step.get("window_index"), step.get("window_handle")))
        iframes.append(str(step.get("iframe_path", "")))
        if step.get("field"):
            fields.append(str(step.get("field")))
        action = normalize_text(step.get("action"))
        if action in {"click", "submit", "enter"}:
            buttons.append(str(step.get("label") or step.get("text") or step.get("css_selector") or step.get("relative_xpath") or ""))
        css = str(step.get("css_selector") or "")
        rel = str(step.get("relative_xpath") or "")
        abs_x = str(step.get("absolute_xpath") or "")
        for selector in (css, rel, abs_x):
            if selector:
                selectors.add(selector)
        candidates = step.get("selector_candidates") or []
        for candidate in candidates:
            try:
                if not bool(candidate.get("unique")) or float(candidate.get("score") or 0.0) < 60:
                    fragile += 1
                wait_conditions.add(str(candidate.get("reason") or ""))
            except Exception:
                continue
        wait_condition = str(step.get("wait_condition") or "")
        if wait_condition:
            wait_conditions.add(wait_condition)

    pages_visited = sorted({normalize_url(url) for url in page_urls if url})
    windows_used = sorted({f"{idx}:{handle}" for idx, handle in windows if handle})
    iframes_used = sorted({iframe for iframe in iframes if iframe})
    fields_used = sorted({normalize_text(field) for field in fields if field})
    buttons_clicked = sorted({normalize_text(button) for button in buttons if button})

    return {
        "process_name": process_name,
        "generated_at": utc_now_iso(),
        "step_count": len(steps),
        "pages_visited": pages_visited,
        "windows_used": windows_used,
        "iframes_used": iframes_used,
        "fields_filled": fields_used,
        "buttons_clicked": buttons_clicked,
        "unique_selectors": sorted(selectors),
        "fragile_selectors": fragile,
        "wait_conditions_suggested": sorted({cond for cond in wait_conditions if cond}),
    }


def _sanitize_raw_event(event: Mapping[str, Any]) -> dict[str, Any]:
    item = dict(event)
    action = normalize_text(item.get("action"))
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
    elif action in {"click", "submit", "focus", "blur", "marker", "checkpoint", "pause", "resume", "modal_visible", "modal_hidden", "element_visible", "element_hidden", "url_changed", "new_window_requested", "alert", "confirm", "prompt", "enter"}:
        item.pop("value", None)
    return sanitize_json_value(item)


@dataclass
class InteractionRecorder:
    process_name: str
    root: Path
    site_recorder: Any
    dom_inventory: DomInventory
    poll_seconds: float = 0.5
    capture_hidden: bool = False
    paused: bool = False
    stopped: bool = False
    raw_events: list[dict[str, Any]] = field(default_factory=list)
    page_snapshots: list[dict[str, Any]] = field(default_factory=list)
    _last_state_by_context: dict[str, dict[str, Any]] = field(default_factory=dict)
    _seen_signatures: set[str] = field(default_factory=set)
    _step_number: int = 0
    _last_page_signature: str = ""

    def __post_init__(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        self.screenshots_dir = self.root / "screenshots"
        self.dom_dir = self.root / "dom"
        self.screenshots_dir.mkdir(parents=True, exist_ok=True)
        self.dom_dir.mkdir(parents=True, exist_ok=True)

    def install(self, driver: Any) -> None:
        self._inject_context(driver)

    def _inject_context(self, driver: Any) -> None:
        try:
            driver.execute_script(DEFAULT_INTERACTION_SCRIPT)
        except Exception:
            pass
        try:
            iframe_elements = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return
        for iframe in iframe_elements:
            try:
                driver.switch_to.frame(iframe)
                self._inject_context(driver)
            except Exception:
                continue
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass

    def _collect_context_events(self, driver: Any, frame_path: str = "top") -> list[dict[str, Any]]:
        events: list[dict[str, Any]] = []
        try:
            raw = driver.execute_script(
                """
                const rec = window.__SOMA_INTERACTION_RECORDER__;
                return rec && rec.flush ? rec.flush() : [];
                """
            )
            for item in raw or []:
                event = dict(item)
                event["iframe_path"] = frame_path
                events.append(_sanitize_raw_event(event))
        except Exception:
            pass

        try:
            iframe_elements = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return events

        for index, iframe in enumerate(iframe_elements):
            child_path = f"{frame_path}/{index}"
            try:
                driver.switch_to.frame(iframe)
                events.extend(self._collect_context_events(driver, child_path))
            except Exception:
                continue
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass
        return events

    def _capture_snapshot(self, driver: Any, reason: str) -> PageSnapshot | None:
        try:
            page_id = f"{self.process_name}_{len(self.page_snapshots) + 1:04d}"
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
            )
            if snapshot.signature in self._seen_signatures:
                return None
            self._seen_signatures.add(snapshot.signature)
            if hasattr(self.site_recorder, "save_page"):
                try:
                    self.site_recorder.save_page(snapshot)
                except Exception:
                    pass
            self.page_snapshots.append(snapshot.to_dict())
            log_kv(log, "Snapshot gravado.", level=logging.INFO, reason=reason, signature=snapshot.signature, page=snapshot.url)
            return snapshot
        except Exception as exc:
            log.debug("Falha a gravar snapshot: %s", exc)
            return None

    def capture_checkpoint(self, driver: Any, label: str) -> dict[str, Any]:
        state = collect_page_state(driver)
        snapshot = self._capture_snapshot(driver, reason=label)
        payload = {
            "action": "checkpoint",
            "label": sanitize_text(label),
            "timestamp": utc_now_iso(),
            "page_url": sanitize_text(state.get("url", "")),
            "page_title": sanitize_text(state.get("title", "")),
            "window_index": 0,
            "window_handle": sanitize_text(state.get("window_handle", "")),
            "iframe_path": sanitize_text(state.get("iframe_path", "top")),
            "before_signature": page_state_signature(state),
            "after_signature": snapshot.signature if snapshot else page_state_signature(state),
        }
        self.raw_events.append(payload)
        return payload

    def record_marker(self, driver: Any, label: str) -> dict[str, Any]:
        state = collect_page_state(driver)
        payload = {
            "action": "marker",
            "label": sanitize_text(label),
            "timestamp": utc_now_iso(),
            "page_url": sanitize_text(state.get("url", "")),
            "page_title": sanitize_text(state.get("title", "")),
            "window_index": 0,
            "window_handle": sanitize_text(state.get("window_handle", "")),
            "iframe_path": sanitize_text(state.get("iframe_path", "top")),
            "before_signature": page_state_signature(state),
            "after_signature": page_state_signature(state),
        }
        self.raw_events.append(payload)
        return payload

    def pause(self, driver: Any | None = None) -> None:
        self.paused = True
        if driver is not None:
            try:
                self.record_marker(driver, "pause")
            except Exception:
                pass

    def resume(self, driver: Any | None = None) -> None:
        self.paused = False
        if driver is not None:
            try:
                self.record_marker(driver, "resume")
            except Exception:
                pass

    def request_stop(self, driver: Any | None = None) -> None:
        self.stopped = True
        if driver is not None:
            try:
                self.record_marker(driver, "stop")
            except Exception:
                pass

    def process_command(self, command: str, driver: Any) -> str:
        cmd = normalize_text(command)
        if not cmd:
            self.capture_checkpoint(driver, "checkpoint")
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
        accepted: list[dict[str, Any]] = []
        if self.paused or self.stopped:
            return accepted
        for event in events:
            item = dict(event)
            state = collect_page_state(driver)
            signature = page_state_signature(state)
            selector_candidates, selector_recommended = build_selector_candidate_payload(driver, item)
            before_state = self._last_state_by_context.get(str(item.get("iframe_path") or "top"), state)
            after_state = dict(state)
            item.setdefault("before_signature", page_state_signature(before_state))
            item.setdefault("after_signature", signature)
            item.setdefault("wait_condition", suggest_wait_condition(before_state, after_state))
            item.setdefault("window_index", 0)
            item.setdefault("window_handle", state.get("window_handle", ""))
            item.setdefault("page_url", state.get("url", item.get("page_url", "")))
            item.setdefault("page_title", state.get("title", item.get("page_title", "")))
            item.setdefault("iframe_path", item.get("iframe_path") or state.get("iframe_path", "top"))
            item.setdefault("selector_candidates", selector_candidates)
            item.setdefault("selector_recommended", selector_recommended)
            item.setdefault("timestamp", utc_now_iso())
            accepted.append(_sanitize_raw_event(item))
            self.raw_events.append(_sanitize_raw_event(item))
            self._last_state_by_context[str(item.get("iframe_path") or "top")] = after_state
            if signature and signature != self._last_page_signature:
                self._last_page_signature = signature
                self._capture_snapshot(driver, reason=normalize_text(item.get("action")) or "event")

        return accepted

    def poll(self, driver: Any) -> list[dict[str, Any]]:
        if self.paused or self.stopped:
            return []
        events = self._collect_context_events(driver)
        if not events:
            state = collect_page_state(driver)
            signature = page_state_signature(state)
            if signature and signature != self._last_page_signature:
                self._last_page_signature = signature
                self._capture_snapshot(driver, reason="dom_changed")
            return []
        return self.ingest(driver, events)

    def finalize(self, driver: Any | None = None) -> dict[str, Any]:
        if driver is not None:
            try:
                self.poll(driver)
            except Exception:
                pass
        workflow = consolidate_events(self.raw_events)
        steps: list[dict[str, Any]] = []
        previous_state: dict[str, Any] = {}
        for idx, event in enumerate(workflow, start=1):
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
            step = event_to_step(event, step_number=idx, before_state=previous_state or state, after_state=state)
            steps.append(step)
            previous_state = state

        workflow_summary = build_workflow_summary(self.process_name, steps)
        elements_used = self._build_elements_used(steps)
        locator_candidates = self._build_locator_candidates_used(steps)

        self._write_json("steps.json", self.raw_events)
        self._write_json("workflow.json", {"process_name": self.process_name, "steps": steps})
        self._write_json("workflow_summary.json", workflow_summary)
        self._write_json("elements_used.json", elements_used)
        self._write_json("locator_candidates_used.json", locator_candidates)

        return {
            "steps": steps,
            "workflow_summary": workflow_summary,
            "elements_used": elements_used,
            "locator_candidates_used": locator_candidates,
        }

    def _write_json(self, name: str, payload: Any) -> None:
        path = self.root / name
        path.write_text(json.dumps(sanitize_json_value(payload), ensure_ascii=False, indent=2), encoding="utf-8")

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
        seen: set[tuple[str, str]] = set()
        for step in steps:
            for candidate in step.get("selector_candidates", []) or []:
                key = (normalize_text(candidate.get("by")), normalize_text(candidate.get("selector")))
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(
                    sanitize_json_value(
                        {
                            "page_url": step.get("page_url", ""),
                            "page_title": step.get("page_title", ""),
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
    "event_to_step",
    "page_state_signature",
    "sanitize_record_value",
    "suggest_wait_condition",
]
