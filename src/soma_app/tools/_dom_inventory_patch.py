from __future__ import annotations

import json
from typing import Any

from selenium.common.exceptions import WebDriverException

from soma_app.automation.dom_inventory import DomInventory

_INTERACTIVE_SELECTOR = (
    "input, select, textarea, button, a, form, table, thead, tbody, tr, td, th, label, iframe, "
    "[role], [contenteditable], [data-testid], [data-test], [data-qa], [data-cy], [data-id], [data-name]"
)


def _collect_context_payload_fixed(self: DomInventory, driver: Any) -> dict[str, Any]:
    """Read the current DOM context using JavaScript that is safe to compile in Chrome."""
    script = r"""
    const absoluteXPath = (el) => {
      if (!el) return '';
      const segments = [];
      let node = el;
      while (node && node.nodeType === Node.ELEMENT_NODE) {
        const tag = (node.tagName || '').toLowerCase();
        if (!tag) break;
        let index = 1;
        let sibling = node.previousElementSibling;
        while (sibling) {
          if ((sibling.tagName || '').toLowerCase() === tag) index += 1;
          sibling = sibling.previousElementSibling;
        }
        segments.unshift(`${tag}[${index}]`);
        node = node.parentElement;
        if (tag === 'html') break;
      }
      return '/' + segments.join('/');
    };

    const stableDataAttrs = (el) => {
      const out = {};
      const allowed = new Set([
        'data-testid',
        'data-test',
        'data-qa',
        'data-cy',
        'data-id',
        'data-name'
      ]);
      Array.from(el.attributes || []).forEach((attr) => {
        if (allowed.has(attr.name)) out[attr.name] = attr.value || '';
      });
      return out;
    };

    const labelsFor = (el) => {
      try {
        if (el.labels && el.labels.length) {
          return Array.from(el.labels)
            .map((label) => (label.innerText || label.textContent || '').trim())
            .filter(Boolean)
            .join(' | ');
        }
      } catch (error) {}
      const id = el.getAttribute('id');
      if (!id) return '';
      try {
        const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
        return label ? (label.innerText || label.textContent || '').trim() : '';
      } catch (error) {
        return '';
      }
    };

    const labelledBy = (el) => {
      const value = el.getAttribute('aria-labelledby') || '';
      if (!value.trim()) return '';
      return value
        .split(/\s+/)
        .map((id) => {
          const referenced = document.getElementById(id);
          return referenced ? (referenced.innerText || referenced.textContent || '').trim() : '';
        })
        .filter(Boolean)
        .join(' | ');
    };

    const nodes = Array.from(document.querySelectorAll(__INTERACTIVE_SELECTOR__));
    const all = Array.from(new Set(nodes));

    return all.map((el, index) => {
      const rect = el.getBoundingClientRect();
      const style = window.getComputedStyle(el);
      const form = el.closest ? el.closest('form') : null;
      const tag = (el.tagName || '').toLowerCase();
      const textAllowed = ['a', 'button', 'label', 'option', 'span', 'th'].includes(tag);
      return {
        index,
        tag,
        id: el.getAttribute('id') || '',
        name: el.getAttribute('name') || '',
        type: el.getAttribute('type') || '',
        class_name: el.getAttribute('class') || '',
        text: textAllowed ? (el.innerText || el.textContent || '').trim().slice(0, 120) : '',
        placeholder: el.getAttribute('placeholder') || '',
        title: el.getAttribute('title') || '',
        role: el.getAttribute('role') || '',
        aria_label: el.getAttribute('aria-label') || '',
        aria_labelledby: el.getAttribute('aria-labelledby') || '',
        data_attrs: stableDataAttrs(el),
        label: labelsFor(el),
        labelled_by: labelledBy(el),
        visible: !!(rect.width || rect.height)
          && style.display !== 'none'
          && style.visibility !== 'hidden'
          && style.opacity !== '0',
        enabled: !el.disabled,
        x: Math.round(rect.x || 0),
        y: Math.round(rect.y || 0),
        width: Math.round(rect.width || 0),
        height: Math.round(rect.height || 0),
        absolute_xpath: absoluteXPath(el),
        in_form: !!form,
        form_text: '',
        frame_count: document.querySelectorAll('iframe').length,
        html: '',
      };
    });
    """
    script = script.replace("__INTERACTIVE_SELECTOR__", json.dumps(_INTERACTIVE_SELECTOR))
    try:
        payload = driver.execute_script(script)
        return {"elements": payload or []}
    except WebDriverException:
        raise
    except Exception as exc:
        self.log.warning("Falha ao ler contexto DOM: %s", exc)
        return {"elements": []}


def apply_dom_inventory_javascript_fix() -> None:
    if getattr(DomInventory, "_soma_javascript_fix_applied", False):
        return
    DomInventory._collect_context_payload = _collect_context_payload_fixed  # type: ignore[method-assign]
    DomInventory._soma_javascript_fix_applied = True  # type: ignore[attr-defined]


__all__ = ["apply_dom_inventory_javascript_fix"]
