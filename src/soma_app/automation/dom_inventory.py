from __future__ import annotations

import hashlib
import json
import logging
import re
import time
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping, Sequence
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By

log = logging.getLogger("soma_app.automation.dom_inventory")

_INTERACTIVE_SELECTOR = (
    "input, select, textarea, button, a, form, table, thead, tbody, tr, td, th, label, iframe, "
    "[role], [contenteditable], [data-testid], [data-test], [data-qa], [data-cy], [data-id], [data-name]"
)

_DATA_ATTR_PRIORITY = (
    "data-testid",
    "data-test",
    "data-qa",
    "data-cy",
    "data-id",
    "data-name",
)

_DANGEROUS_TEXT_RE = re.compile(
    r"(?i)\b("
    r"salvar|gravar|confirmar|eliminar|apagar|pagar|pagamento|realizar pagamento|inserir baixa|"
    r"submeter|processar|executar|cancelar documento|excluir|remove|delete|submit|save|confirm|"
    r"cancel|pay|payment"
    r")\b"
)

_DANGEROUS_SELECTOR_RE = re.compile(
    r"(?i)(button\[type=['\"]submit['\"]\]|input\[type=['\"]submit['\"]|\.btn-danger|\.danger|"
    r"\[class[^\]]*(?:danger|dangerous|delete|remove|submit|save|confirm|cancel)[^\]]*\])"
)

_SENSITIVE_VALUE_RE = re.compile(
    r"(?i)("
    r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+|"
    r"\b(?:\d[ -]*?){11,19}\b|"
    r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b|"
    r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b|"
    r"(?:r\$|€|\$)\s?\d[\d\.\,]*|"
    r"\b\d{4,}\b"
    r")"
)

_DYNAMIC_CLASS_RE = re.compile(
    r"(?i)(?:^|[\-_])(?:\d{3,}|[a-f0-9]{8,}|[a-f0-9]{12,})(?:$|[\-_])"
)

_POSITIONAL_XPATH_RE = re.compile(r"\[[0-9]+\]")


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _strip_accents(value: str) -> str:
    try:
        import unicodedata

        normalized = unicodedata.normalize("NFKD", value)
        return "".join(ch for ch in normalized if not unicodedata.combining(ch))
    except Exception:
        return value


def normalize_text(value: Any) -> str:
    raw = "" if value is None else str(value)
    return " ".join(_strip_accents(raw).split()).strip().lower()


def sanitize_text(value: Any, *, max_len: int = 240) -> str:
    raw = "" if value is None else str(value)
    text = " ".join(raw.split()).strip()
    if not text:
        return ""
    if _SENSITIVE_VALUE_RE.search(text):
        return "[redacted]"
    if len(text) > max_len:
        return text[:max_len] + "…"
    return text


def redact_sensitive_html(html_text: str) -> str:
    if not html_text:
        return ""

    text = html_text
    text = re.sub(r"(?is)<script\b.*?</script>", "", text)
    text = re.sub(r"(?is)<style\b.*?</style>", "", text)
    text = re.sub(r"(?is)<!--.*?-->", "", text)

    for attr in ("value", "headers", "cookie", "cookies", "token", "auth", "authorization", "password", "secret"):
        text = re.sub(
            rf'(?i)\s{re.escape(attr)}=(["\']).*?\1',
            f' {attr}=""',
            text,
        )

    text = re.sub(
        r"(?is)(<textarea\b[^>]*>)(.*?)(</textarea>)",
        lambda m: m.group(1) + "[redacted]" + m.group(3),
        text,
    )

    def _redact_text_nodes(match: re.Match[str]) -> str:
        snippet = match.group(0)
        if _SENSITIVE_VALUE_RE.search(snippet):
            return "[redacted]"
        return snippet

    text = re.sub(r"[^<>{}\[\]\n]{3,}", _redact_text_nodes, text)
    return text


def sanitize_json_value(value: Any) -> Any:
    if isinstance(value, str):
        return sanitize_text(value)
    if isinstance(value, list):
        return [sanitize_json_value(item) for item in value]
    if isinstance(value, tuple):
        return [sanitize_json_value(item) for item in value]
    if isinstance(value, dict):
        return {str(k): sanitize_json_value(v) for k, v in value.items()}
    return value


def normalize_url(url: str) -> str:
    raw = (url or "").strip()
    if not raw:
        return ""

    parts = urlsplit(raw)
    scheme = parts.scheme.lower()
    hostname = (parts.hostname or "").lower()
    if not scheme or not hostname:
        return raw.rstrip("#")

    port = parts.port
    default_port = (scheme == "http" and port == 80) or (scheme == "https" and port == 443)
    if port and not default_port:
        netloc = f"{hostname}:{port}"
    else:
        netloc = hostname

    path = re.sub(r"/{2,}", "/", parts.path or "/")
    if path != "/" and path.endswith("/"):
        path = path.rstrip("/")
    if not path:
        path = "/"

    query = urlencode(sorted(parse_qsl(parts.query, keep_blank_values=True)))
    return urlunsplit((scheme, netloc, path, query, ""))


def page_signature(normalized_url: str, title: str, structure_hash: str, interactive_hash: str) -> str:
    payload = "|".join((normalized_url or "", normalize_text(title), structure_hash or "", interactive_hash or ""))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _css_attr(value: Any) -> str:
    raw = "" if value is None else str(value)
    raw = raw.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{raw}"'


def _xpath_literal(value: Any) -> str:
    raw = "" if value is None else str(value)
    if "'" not in raw:
        return f"'{raw}'"
    if '"' not in raw:
        return f'"{raw}"'
    parts = raw.split("'")
    return "concat(" + ", \"'\", ".join(f"'{part}'" for part in parts) + ")"


def _stable_class_tokens(value: Any) -> list[str]:
    raw = "" if value is None else str(value)
    tokens: list[str] = []
    for token in raw.split():
        cleaned = token.strip()
        if not cleaned or len(cleaned) < 2:
            continue
        if _DYNAMIC_CLASS_RE.search(cleaned):
            continue
        if cleaned.lower() in {"active", "selected", "open", "show", "hidden"}:
            continue
        tokens.append(cleaned)
    return tokens[:3]


def _selector_contains_position(selector: str) -> bool:
    return bool(_POSITIONAL_XPATH_RE.search(selector) or re.search(r"\bdiv\[\d+\]", selector))


def is_dangerous_text(text: Any) -> bool:
    return bool(_DANGEROUS_TEXT_RE.search(normalize_text(text)))


def is_dangerous_selector(selector: str) -> bool:
    if not selector:
        return False
    return bool(_DANGEROUS_SELECTOR_RE.search(selector))


def is_safe_auto_click(
    *,
    tag: str,
    text: str,
    selector: str,
    in_form: bool,
    allowed_selector: bool = False,
) -> bool:
    if is_dangerous_text(text) or is_dangerous_selector(selector):
        return False
    if in_form and not allowed_selector:
        return False
    if tag.lower() == "input" and normalize_text(selector).find("type=\"submit\"") >= 0:
        return False
    return True


@dataclass
class SelectorCandidate:
    strategy: str
    by: str
    selector: str
    base_score: float
    reason: str
    found_count: int = 0
    unique: bool = False
    score: float = 0.0

    def __post_init__(self) -> None:
        if not self.score:
            self.score = float(self.base_score)


@dataclass
class ElementSnapshot:
    element_id: str
    tag: str
    id: str
    name: str
    type: str
    class_name: str
    text: str
    placeholder: str
    title: str
    role: str
    aria_label: str
    aria_labelledby: str
    data_attrs: dict[str, str]
    label: str
    url: str
    page_title: str
    window_index: int
    window_handle: str
    iframe_path: str
    visible: bool
    enabled: bool
    x: int
    y: int
    width: int
    height: int
    relative_xpath: str
    css_selector: str
    absolute_xpath: str
    selector_confidence: float
    selector_candidates: list[SelectorCandidate] = field(default_factory=list)
    timestamp: str = ""
    in_form: bool = False
    form_label: str = ""

    @property
    def interactive(self) -> bool:
        return self.tag.lower() in {"input", "select", "textarea", "button", "a", "form", "table", "iframe"} or bool(self.role or self.aria_label)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["selector_candidates"] = [asdict(item) for item in self.selector_candidates]
        return payload


@dataclass
class FrameSnapshot:
    frame_path: str
    iframe_index: int
    iframe_locator: str
    url: str
    title: str
    html_path: str
    html_hash: str
    html: str = ""


@dataclass
class PageSnapshot:
    page_id: str
    url: str
    normalized_url: str
    title: str
    window_index: int
    window_handle: str
    iframe_path: str
    capture_timestamp: str
    structure_hash: str
    interactive_hash: str
    signature: str
    element_count: int
    interactive_count: int
    visible_count: int
    hidden_count: int
    screenshot_path: str
    dom_html_path: str
    frame_count: int
    dom_html: str = ""
    elements: list[ElementSnapshot] = field(default_factory=list)
    frames: list[FrameSnapshot] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["elements"] = [item.to_dict() for item in self.elements]
        payload["frames"] = [{k: v for k, v in asdict(item).items() if k != "html"} for item in self.frames]
        payload.pop("dom_html", None)
        return payload


@dataclass
class SelectorValidation:
    by: str
    selector: str
    count: int
    unique: bool
    score: float
    reason: str


class CaptureTracker:
    def __init__(self) -> None:
        self._seen: set[str] = set()

    def should_capture(self, signature: str) -> bool:
        signature = (signature or "").strip()
        if not signature:
            return True
        if signature in self._seen:
            return False
        self._seen.add(signature)
        return True

    def remember(self, signature: str) -> None:
        signature = (signature or "").strip()
        if signature:
            self._seen.add(signature)


def css_selector_for_element(element: Mapping[str, Any]) -> str:
    tag = (element.get("tag") or "*").strip().lower() or "*"
    for attr in _DATA_ATTR_PRIORITY:
        data_attrs = element.get("data_attrs") or {}
        if isinstance(data_attrs, Mapping):
            value = sanitize_text(data_attrs.get(attr))
            if value:
                return f'{tag}[{attr}={_css_attr(value)}]' if tag != "*" else f'[{attr}={_css_attr(value)}]'

    element_id = sanitize_text(element.get("id"))
    if element_id:
        return f'{tag}[id={_css_attr(element_id)}]' if tag != "*" else f'[id={_css_attr(element_id)}]'

    name = sanitize_text(element.get("name"))
    if name:
        return f'{tag}[name={_css_attr(name)}]' if tag != "*" else f'[name={_css_attr(name)}]'

    aria_label = sanitize_text(element.get("aria_label"))
    if aria_label:
        return f'{tag}[aria-label={_css_attr(aria_label)}]' if tag != "*" else f'[aria-label={_css_attr(aria_label)}]'

    role = sanitize_text(element.get("role"))
    if role:
        return f'{tag}[role={_css_attr(role)}]' if tag != "*" else f'[role={_css_attr(role)}]'

    classes = _stable_class_tokens(element.get("class_name"))
    if classes:
        class_selector = "".join(f".{token}" for token in classes)
        return f"{tag}{class_selector}" if tag != "*" else class_selector

    placeholder = sanitize_text(element.get("placeholder"))
    if placeholder and tag in {"input", "textarea", "select"}:
        return f'{tag}[placeholder={_css_attr(placeholder)}]'

    title = sanitize_text(element.get("title"))
    if title:
        return f'{tag}[title={_css_attr(title)}]' if tag != "*" else f'[title={_css_attr(title)}]'

    return ""


def xpath_relative_for_element(element: Mapping[str, Any], *, label_text: str = "") -> str:
    tag = (element.get("tag") or "*").strip().lower() or "*"
    for attr in _DATA_ATTR_PRIORITY:
        data_attrs = element.get("data_attrs") or {}
        if isinstance(data_attrs, Mapping):
            value = sanitize_text(data_attrs.get(attr))
            if value:
                return f"//*[@{attr}={_xpath_literal(value)}]"

    element_id = sanitize_text(element.get("id"))
    if element_id:
        return f"//*[@id={_xpath_literal(element_id)}]"

    name = sanitize_text(element.get("name"))
    if name:
        return f"//{tag}[@name={_xpath_literal(name)}]" if tag != "*" else f"//*[@name={_xpath_literal(name)}]"

    aria_label = sanitize_text(element.get("aria_label"))
    if aria_label:
        return f"//{tag}[@aria-label={_xpath_literal(aria_label)}]" if tag != "*" else f"//*[@aria-label={_xpath_literal(aria_label)}]"

    role = sanitize_text(element.get("role"))
    if role:
        return f"//{tag}[@role={_xpath_literal(role)}]" if tag != "*" else f"//*[@role={_xpath_literal(role)}]"

    placeholder = sanitize_text(element.get("placeholder"))
    if placeholder and tag in {"input", "textarea", "select"}:
        return f"//{tag}[@placeholder={_xpath_literal(placeholder)}]"

    title = sanitize_text(element.get("title"))
    if title:
        return f"//{tag}[@title={_xpath_literal(title)}]" if tag != "*" else f"//*[@title={_xpath_literal(title)}]"

    if label_text and tag in {"input", "select", "textarea"}:
        label = sanitize_text(label_text)
        if label:
            return f"//label[contains(normalize-space(.), {_xpath_literal(label)})]/following::{tag}[1]"

    text = sanitize_text(element.get("text"))
    if text and tag in {"button", "a", "label", "span"}:
        return f"//{tag}[contains(normalize-space(.), {_xpath_literal(text[:80])})]"

    classes = _stable_class_tokens(element.get("class_name"))
    if classes:
        class_expr = " and ".join(
            f"contains(concat(' ', normalize-space(@class), ' '), {_xpath_literal(' ' + token + ' ')})" for token in classes
        )
        return f"//{tag}[{class_expr}]"

    return ""


def absolute_xpath_from_segments(segments: Sequence[Sequence[Any]]) -> str:
    parts: list[str] = []
    for item in segments:
        if not item:
            continue
        tag = str(item[0]).strip().lower() or "*"
        index = 1
        if len(item) > 1:
            try:
                index = int(item[1])
            except Exception:
                index = 1
        parts.append(f"{tag}[{index}]")
    return "/" + "/".join(parts)


def selector_candidates_for_element(element: Mapping[str, Any], *, label_text: str = "") -> list[SelectorCandidate]:
    tag = (element.get("tag") or "*").strip().lower() or "*"
    text = sanitize_text(element.get("text"))
    role = sanitize_text(element.get("role"))
    aria_label = sanitize_text(element.get("aria_label"))
    element_id = sanitize_text(element.get("id"))
    name = sanitize_text(element.get("name"))
    classes = _stable_class_tokens(element.get("class_name"))
    placeholder = sanitize_text(element.get("placeholder"))
    title = sanitize_text(element.get("title"))
    data_attrs = element.get("data_attrs") if isinstance(element.get("data_attrs"), Mapping) else {}

    candidates: list[SelectorCandidate] = []

    def add(strategy: str, by: str, selector: str, score: float, reason: str) -> None:
        selector = (selector or "").strip()
        if not selector:
            return
        if any(item.selector == selector and item.by == by for item in candidates):
            return
        candidates.append(SelectorCandidate(strategy=strategy, by=by, selector=selector, base_score=score, reason=reason))

    for attr in _DATA_ATTR_PRIORITY:
        value = sanitize_text(data_attrs.get(attr)) if isinstance(data_attrs, Mapping) else ""
        if value:
            add(
                strategy=attr,
                by=By.CSS_SELECTOR,
                selector=f'[{attr}={_css_attr(value)}]',
                score=100.0,
                reason=f"atributo data estável ({attr})",
            )
            add(
                strategy=attr,
                by=By.XPATH,
                selector=f"//*[@{attr}={_xpath_literal(value)}]",
                score=98.0,
                reason=f"atributo data estável ({attr})",
            )
            break

    if element_id:
        add("id", By.CSS_SELECTOR, f'[id={_css_attr(element_id)}]', 95.0, "id único ou estável")
        add("id", By.XPATH, f"//*[@id={_xpath_literal(element_id)}]", 94.0, "id único ou estável")

    if name:
        add("name", By.CSS_SELECTOR, f'{tag}[name={_css_attr(name)}]' if tag != "*" else f'[name={_css_attr(name)}]', 90.0, "name estável")
        add("name", By.XPATH, f"//{tag}[@name={_xpath_literal(name)}]" if tag != "*" else f"//*[@name={_xpath_literal(name)}]", 88.0, "name estável")

    if aria_label:
        add("aria-label", By.CSS_SELECTOR, f'{tag}[aria-label={_css_attr(aria_label)}]' if tag != "*" else f'[aria-label={_css_attr(aria_label)}]', 85.0, "aria-label estável")
        add("aria-label", By.XPATH, f"//{tag}[@aria-label={_xpath_literal(aria_label)}]" if tag != "*" else f"//*[@aria-label={_xpath_literal(aria_label)}]", 84.0, "aria-label estável")

    if role:
        add("role", By.CSS_SELECTOR, f'{tag}[role={_css_attr(role)}]' if tag != "*" else f'[role={_css_attr(role)}]', 78.0, "role semântico")
        add("role", By.XPATH, f"//{tag}[@role={_xpath_literal(role)}]" if tag != "*" else f"//*[@role={_xpath_literal(role)}]", 77.0, "role semântico")

    if placeholder and tag in {"input", "textarea", "select"}:
        add(
            "placeholder",
            By.CSS_SELECTOR,
            f'{tag}[placeholder={_css_attr(placeholder)}]',
            72.0,
            "placeholder descritivo",
        )
        add(
            "placeholder",
            By.XPATH,
            f"//{tag}[@placeholder={_xpath_literal(placeholder)}]",
            71.0,
            "placeholder descritivo",
        )

    if title:
        add("title", By.CSS_SELECTOR, f'{tag}[title={_css_attr(title)}]' if tag != "*" else f'[title={_css_attr(title)}]', 68.0, "title descritivo")
        add("title", By.XPATH, f"//{tag}[@title={_xpath_literal(title)}]" if tag != "*" else f"//*[@title={_xpath_literal(title)}]", 67.0, "title descritivo")

    if label_text and tag in {"input", "select", "textarea"}:
        label = sanitize_text(label_text)
        if label:
            add(
                "label",
                By.XPATH,
                f"//label[contains(normalize-space(.), {_xpath_literal(label)})]/following::{tag}[1]",
                80.0,
                "associação com label",
            )

    if text and tag in {"button", "a"}:
        add("text", By.XPATH, f"//{tag}[contains(normalize-space(.), {_xpath_literal(text[:80])})]", 65.0, "texto visível estável")

    if classes:
        css = tag + "".join(f".{token}" for token in classes)
        add("class", By.CSS_SELECTOR, css, 60.0, "classes estáveis")
        class_expr = " and ".join(
            f"contains(concat(' ', normalize-space(@class), ' '), {_xpath_literal(' ' + token + ' ')})" for token in classes
        )
        add("class", By.XPATH, f"//{tag}[{class_expr}]", 58.0, "classes estáveis")

    if tag in {"input", "select", "textarea", "button", "a", "form", "table", "thead", "tbody", "tr", "td", "th", "label", "iframe"}:
        add("tag", By.XPATH, f"//{tag}", 30.0, "fallback por tag")

    return candidates


def score_selector_candidate(
    candidate: SelectorCandidate,
    *,
    count: int,
    stable_bonus: float = 0.0,
) -> float:
    score = float(candidate.base_score)

    if count == 1:
        score += 18.0
    elif count == 0:
        score -= 42.0
    elif count > 1:
        score -= min(40.0, 8.0 * count)

    if candidate.by == By.XPATH and candidate.selector.startswith("/"):
        score -= 10.0
    if _selector_contains_position(candidate.selector):
        score -= 20.0
    if candidate.strategy == "class" and _DYNAMIC_CLASS_RE.search(candidate.selector):
        score -= 15.0

    score += stable_bonus
    return max(0.0, min(100.0, score))


def validate_selector(driver: Any, candidate: SelectorCandidate) -> SelectorValidation:
    try:
        elements = driver.find_elements(candidate.by, candidate.selector)
        count = len(elements)
        unique = count == 1
        score = score_selector_candidate(candidate, count=count)
        reason = candidate.reason
        if count == 0:
            reason = f"{reason}; não encontrou elementos"
        elif count == 1:
            reason = f"{reason}; único"
        else:
            reason = f"{reason}; duplicado ({count})"
        return SelectorValidation(
            by=candidate.by,
            selector=candidate.selector,
            count=count,
            unique=unique,
            score=score,
            reason=reason,
        )
    except Exception as exc:
        return SelectorValidation(
            by=candidate.by,
            selector=candidate.selector,
            count=0,
            unique=False,
            score=0.0,
            reason=f"{candidate.reason}; erro ao validar: {exc}",
        )


def best_selector_candidate(driver: Any, candidates: Iterable[SelectorCandidate]) -> SelectorCandidate | None:
    best: SelectorCandidate | None = None
    for candidate in candidates:
        validation = validate_selector(driver, candidate)
        candidate.found_count = validation.count
        candidate.unique = validation.unique
        candidate.score = validation.score
        candidate.reason = validation.reason
        if best is None:
            best = candidate
            continue
        if candidate.unique and not best.unique:
            best = candidate
            continue
        if candidate.unique == best.unique and candidate.score > best.score:
            best = candidate
    return best


def build_absolute_xpath_from_segments(segments: Sequence[Sequence[Any]]) -> str:
    return absolute_xpath_from_segments(segments)


class DomInventory:
    def __init__(self, *, capture_hidden: bool = False, max_frame_depth: int = 1, logger: logging.Logger | None = None):
        self.capture_hidden = bool(capture_hidden)
        self.max_frame_depth = max(0, int(max_frame_depth))
        self.log = logger or log

    def _wait_ready(self, driver: Any, timeout_seconds: int = 20, stable_seconds: float = 0.75) -> None:
        end = time.time() + timeout_seconds
        last_sig: str | None = None
        stable_since = 0.0
        while time.time() < end:
            try:
                ready = driver.execute_script("return document.readyState")
                body = driver.execute_script(
                    """
                    const body = document.body || document.documentElement;
                    const text = body ? (body.innerText || body.textContent || '') : '';
                    const interactive = document.querySelectorAll('input, select, textarea, button, a, [role], [contenteditable], iframe').length;
                    return [document.title || '', text.length, interactive, document.querySelectorAll('iframe').length].join('|');
                    """
                )
                signature = f"{ready}|{body}"
                now = time.time()
                if ready == "complete" and signature == last_sig:
                    if stable_since and now - stable_since >= stable_seconds:
                        return
                else:
                    stable_since = now
                    last_sig = signature
            except Exception:
                pass
            time.sleep(0.25)
        self.log.debug("DOM não estabilizou a tempo")

    def _collect_context_payload(self, driver: Any) -> dict[str, Any]:
        script = """
        const isHidden = (el) => {
          const rect = el.getBoundingClientRect();
          const style = window.getComputedStyle(el);
          return !(rect.width || rect.height) || style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0';
        };

        const escapeXPathLiteral = (value) => {
          const s = String(value ?? '');
          if (!s.includes("'")) return `'${s}'`;
          if (!s.includes('"')) return `"${s}"`;
          return 'concat(' + s.split("'").map(part => `'${part}'`).join(", \"'\", ") + ')';
        };

        const absoluteXPath = (el) => {
          if (!el) return '';
          const id = el.getAttribute('id');
          if (id) return `//*[@id=${escapeXPathLiteral(id)}]`;
          const segments = [];
          let node = el;
          while (node && node.nodeType === Node.ELEMENT_NODE) {
            const tag = node.tagName.toLowerCase();
            let index = 1;
            let sib = node.previousElementSibling;
            while (sib) {
              if (sib.tagName.toLowerCase() === tag) index += 1;
              sib = sib.previousElementSibling;
            }
            segments.unshift(`${tag}[${index}]`);
            node = node.parentElement;
            if (tag === 'html') break;
          }
          return '/' + segments.join('/');
        };

        const stableDataAttrs = (el) => {
          const out = {};
          Array.from(el.attributes || []).forEach(attr => {
            if (attr.name && attr.name.startsWith('data-')) out[attr.name] = attr.value || '';
          });
          return out;
        };

        const labelsFor = (el) => {
          try {
            if (el.labels && el.labels.length) {
              return Array.from(el.labels).map(x => (x.innerText || x.textContent || '').trim()).filter(Boolean).join(' | ');
            }
          } catch (e) {}
          const id = el.getAttribute('id');
          if (!id) return '';
          const label = document.querySelector(`label[for="${CSS.escape(id)}"]`);
          return label ? (label.innerText || label.textContent || '').trim() : '';
        };

        const labelledBy = (el) => {
          const value = el.getAttribute('aria-labelledby') || '';
          if (!value.trim()) return '';
          return value.split(/\\s+/).map(id => {
            const ref = document.getElementById(id);
            return ref ? (ref.innerText || ref.textContent || '').trim() : '';
          }).filter(Boolean).join(' | ');
        };

        const nodes = Array.from(document.querySelectorAll(__INTERACTIVE_SELECTOR__));
        const dataNodes = Array.from(document.querySelectorAll('*')).filter(el => Array.from(el.attributes || []).some(attr => attr.name.startsWith('data-')));
        const all = Array.from(new Set([...nodes, ...dataNodes]));

        return all.map((el, index) => {
          const rect = el.getBoundingClientRect();
          const style = window.getComputedStyle(el);
          const form = el.closest ? el.closest('form') : null;
          return {
            index,
            tag: (el.tagName || '').toLowerCase(),
            id: el.getAttribute('id') || '',
            name: el.getAttribute('name') || '',
            type: el.getAttribute('type') || '',
            class_name: el.getAttribute('class') || '',
            text: (el.innerText || el.textContent || '').trim(),
            placeholder: el.getAttribute('placeholder') || '',
            title: el.getAttribute('title') || '',
            role: el.getAttribute('role') || '',
            aria_label: el.getAttribute('aria-label') || '',
            aria_labelledby: el.getAttribute('aria-labelledby') || '',
            data_attrs: stableDataAttrs(el),
            label: labelsFor(el),
            labelled_by: labelledBy(el),
            visible: !!(rect.width || rect.height) && style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0',
            enabled: !el.disabled,
            x: Math.round(rect.x || 0),
            y: Math.round(rect.y || 0),
            width: Math.round(rect.width || 0),
            height: Math.round(rect.height || 0),
            absolute_xpath: absoluteXPath(el),
            in_form: !!form,
            form_text: form ? (form.innerText || form.textContent || '').trim() : '',
            frame_count: document.querySelectorAll('iframe').length,
            html: el.outerHTML || '',
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

    def _collect_frames(self, driver: Any, frame_path: str, depth: int, page_url: str, page_title: str, out_frames: list[FrameSnapshot], out_html: list[tuple[str, str]]) -> None:
        try:
            iframe_elements = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return

        if depth >= self.max_frame_depth:
            return

        for index, iframe_el in enumerate(iframe_elements):
            locator = ""
            try:
                locator = sanitize_text(iframe_el.get_attribute("id") or iframe_el.get_attribute("name") or iframe_el.get_attribute("src") or f"iframe[{index}]")
            except Exception:
                locator = f"iframe[{index}]"

            child_path = f"{frame_path}/{index}"
            try:
                driver.switch_to.frame(iframe_el)
                self._wait_ready(driver, timeout_seconds=10, stable_seconds=0.5)
                child_url = ""
                child_title = ""
                try:
                    child_url = normalize_url(driver.current_url)
                except Exception:
                    child_url = page_url
                try:
                    child_title = sanitize_text(driver.title)
                except Exception:
                    child_title = page_title

                html_text = ""
                try:
                    html_text = driver.execute_script("return document.documentElement ? document.documentElement.outerHTML : document.body.outerHTML || ''") or ""
                except Exception:
                    html_text = ""
                html_text = redact_sensitive_html(html_text)
                html_hash = hashlib.sha256(html_text.encode("utf-8")).hexdigest()
                out_html.append((child_path, html_text))
                out_frames.append(
                    FrameSnapshot(
                        frame_path=child_path,
                        iframe_index=index,
                        iframe_locator=locator,
                        url=child_url,
                        title=child_title,
                        html_path="",
                        html_hash=html_hash,
                        html=html_text,
                    )
                )
                self._collect_frames(driver, child_path, depth + 1, child_url, child_title, out_frames, out_html)
            except Exception as exc:
                self.log.debug("Ignorando iframe inacessível: %s", exc)
            finally:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    pass

    def capture(self, driver: Any, *, page_id: str, screenshot_path: str = "", dom_dir: str = "") -> PageSnapshot:
        self._wait_ready(driver)

        current_url = getattr(driver, "current_url", "") or ""
        title = getattr(driver, "title", "") or ""
        normalized = normalize_url(current_url)
        window_handles = list(getattr(driver, "window_handles", []) or [])
        try:
            window_index = window_handles.index(driver.current_window_handle)
        except Exception:
            window_index = 0
        window_handle = getattr(driver, "current_window_handle", "") or ""

        html_text = ""
        try:
            html_text = driver.execute_script("return document.documentElement ? document.documentElement.outerHTML : document.body.outerHTML || ''") or ""
        except Exception:
            html_text = ""
        html_text = redact_sensitive_html(html_text)
        structure_hash = hashlib.sha256(html_text.encode("utf-8")).hexdigest()

        context = self._collect_context_payload(driver)
        raw_elements = context.get("elements") or []

        elements: list[ElementSnapshot] = []
        interactive_signature_bits: list[str] = []
        visible_count = 0
        hidden_count = 0

        for idx, raw in enumerate(raw_elements):
            if not isinstance(raw, Mapping):
                continue
            tag = sanitize_text(raw.get("tag")).lower()
            visible = bool(raw.get("visible"))
            if not visible:
                hidden_count += 1
            else:
                visible_count += 1

            text = sanitize_text(raw.get("text"))
            label = sanitize_text(raw.get("label") or raw.get("labelled_by"))
            data_attrs = raw.get("data_attrs") if isinstance(raw.get("data_attrs"), Mapping) else {}
            element = {
                "tag": tag,
                "id": sanitize_text(raw.get("id")),
                "name": sanitize_text(raw.get("name")),
                "type": sanitize_text(raw.get("type")),
                "class_name": sanitize_text(raw.get("class_name")),
                "text": text,
                "placeholder": sanitize_text(raw.get("placeholder")),
                "title": sanitize_text(raw.get("title")),
                "role": sanitize_text(raw.get("role")),
                "aria_label": sanitize_text(raw.get("aria_label")),
                "data_attrs": {str(k): sanitize_text(v) for k, v in dict(data_attrs).items()},
                "label_text": label,
            }
            css_selector = css_selector_for_element(element)
            xpath_relative = xpath_relative_for_element(element, label_text=label)
            candidates = selector_candidates_for_element(element, label_text=label)
            best = best_selector_candidate(driver, candidates)

            if best is not None:
                interactive_signature_bits.append(f"{best.by}:{best.selector}:{best.score:.1f}")
            else:
                interactive_signature_bits.append(f"{tag}:{text[:40]}")

            frame_path = "top"
            absolute_xpath = sanitize_text(raw.get("absolute_xpath"))
            if not absolute_xpath:
                absolute_xpath = xpath_relative or ""

            snapshot = ElementSnapshot(
                element_id=f"{page_id}-{idx}",
                tag=tag,
                id=element["id"],
                name=element["name"],
                type=element["type"],
                class_name=element["class_name"],
                text=text,
                placeholder=element["placeholder"],
                title=element["title"],
                role=element["role"],
                aria_label=element["aria_label"],
                aria_labelledby=sanitize_text(raw.get("aria_labelledby")),
                data_attrs=element["data_attrs"],
                label=label,
                url=normalized,
                page_title=sanitize_text(title),
                window_index=window_index,
                window_handle=window_handle,
                iframe_path=frame_path,
                visible=visible,
                enabled=bool(raw.get("enabled", True)),
                x=int(raw.get("x") or 0),
                y=int(raw.get("y") or 0),
                width=int(raw.get("width") or 0),
                height=int(raw.get("height") or 0),
                relative_xpath=xpath_relative,
                css_selector=css_selector,
                absolute_xpath=absolute_xpath,
                selector_confidence=float(best.score if best is not None else 0.0),
                selector_candidates=candidates,
                timestamp=utc_now_iso(),
                in_form=bool(raw.get("in_form")),
                form_label=sanitize_text(raw.get("form_text")),
            )

            if not self.capture_hidden and not visible and snapshot.tag not in {"iframe"}:
                continue
            elements.append(snapshot)

        frames: list[FrameSnapshot] = []
        html_snapshots: list[tuple[str, str]] = [("top", html_text)]
        self._collect_frames(driver, "top", 0, normalized, sanitize_text(title), frames, html_snapshots)

        interactive_hash = hashlib.sha256("\n".join(interactive_signature_bits).encode("utf-8")).hexdigest()
        signature = page_signature(normalized, title, structure_hash, interactive_hash)

        for idx, frame in enumerate(frames):
            if dom_dir:
                try:
                    frame_path = f"{page_id}__{frame.frame_path.replace('/', '__')}"
                    frame_file = f"{frame_path}.html"
                    full_path = f"{dom_dir}/{frame_file}"
                    frame.html_path = full_path
                except Exception:
                    pass

        dom_html_path = ""
        if dom_dir:
            dom_html_path = f"{dom_dir}/{page_id}.html"

        return PageSnapshot(
            page_id=page_id,
            url=current_url,
            normalized_url=normalized,
            title=sanitize_text(title),
            window_index=window_index,
            window_handle=window_handle,
            iframe_path="top",
            capture_timestamp=utc_now_iso(),
            structure_hash=structure_hash,
            interactive_hash=interactive_hash,
            signature=signature,
            element_count=len(elements),
            interactive_count=sum(1 for item in elements if item.interactive),
            visible_count=visible_count,
            hidden_count=hidden_count,
            screenshot_path=screenshot_path,
            dom_html_path=dom_html_path,
            frame_count=len(frames),
            dom_html=html_text,
            elements=elements,
            frames=frames,
        )


__all__ = [
    "CaptureTracker",
    "DomInventory",
    "ElementSnapshot",
    "FrameSnapshot",
    "PageSnapshot",
    "SelectorCandidate",
    "SelectorValidation",
    "absolute_xpath_from_segments",
    "best_selector_candidate",
    "build_absolute_xpath_from_segments",
    "css_selector_for_element",
    "is_dangerous_selector",
    "is_dangerous_text",
    "is_safe_auto_click",
    "normalize_text",
    "normalize_url",
    "page_signature",
    "redact_sensitive_html",
    "sanitize_json_value",
    "sanitize_text",
    "score_selector_candidate",
    "selector_candidates_for_element",
    "validate_selector",
    "xpath_relative_for_element",
]
