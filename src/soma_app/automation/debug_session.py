from __future__ import annotations

import datetime as dt
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from selenium.webdriver.common.by import By

from soma_app.infra.env import env_bool, env_str

log = logging.getLogger("soma_app.automation.debug_session")


class GuidedDebugAbort(RuntimeError):
    """Interrompe o fluxo de debug guiado sem encerrar o browser."""


@dataclass(frozen=True)
class DebugCommandResult:
    command: str
    output: str


class GuidedDebugSession:
    def __init__(self, actions: Any, settings: Any | None = None) -> None:
        self.actions = actions
        self.driver = getattr(actions, "driver", None)
        self.enabled = env_bool("DEBUG_STEP_MODE", default=False)
        self.row_filter = self._parse_row_filter(env_str("DEBUG_ROW", "")) if self.enabled else None
        self._secret = str(getattr(settings, "site_password", "") or env_str("SITE_PASSWORD", ""))
        self._counter = 0
        self._observer_installed = False
        self._quit_requested = False
        self._log_path = self._open_log_file(settings)

    @staticmethod
    def _parse_row_filter(raw: str) -> int | None:
        value = (raw or "").strip()
        if not value:
            return None
        try:
            parsed = int(value)
        except ValueError as exc:
            raise ValueError(f"DEBUG_ROW inválido: {value!r}. Use um inteiro positivo ou deixe vazio.") from exc
        if parsed <= 0:
            raise ValueError("DEBUG_ROW deve ser > 0 quando informado.")
        return parsed

    def _open_log_file(self, settings: Any | None) -> Path:
        log_dir = Path(getattr(settings, "log_dir", "logs") or "logs")
        log_dir.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        return log_dir / f"soma_debug_session_{stamp}.log"

    def enabled_for_row(self, row_number: int | None) -> bool:
        if not self.enabled:
            return False
        if self.row_filter is None:
            return True
        return int(row_number or 0) == self.row_filter

    def _write_log(self, text: str) -> None:
        stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        line = f"{stamp} | {text}"
        try:
            self._log_path.parent.mkdir(parents=True, exist_ok=True)
            with self._log_path.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except Exception:
            log.exception("Falha ao gravar log de debug guiado.")

    def _emit(self, text: str) -> None:
        print(text)
        self._write_log(text)

    def _ensure_observer(self) -> None:
        if not self.enabled or self._observer_installed or self.driver is None:
            return
        script = r"""
(function () {
  if (window.__SOMA_DEBUG_MUTATION_OBSERVER_INSTALLED) {
    return "already";
  }
  window.__SOMA_DEBUG_EVENTS = window.__SOMA_DEBUG_EVENTS || [];
  window.__SOMA_DEBUG_LAST_URL = window.location.href;

  function now() {
    return new Date().toISOString();
  }

  function shortText(node) {
    try {
      return (node && node.textContent ? node.textContent : "").replace(/\s+/g, " ").trim().slice(0, 240);
    } catch (err) {
      return "";
    }
  }

  function shortHtml(node) {
    try {
      var html = node && node.outerHTML ? node.outerHTML : "";
      html = html.replace(/\s+/g, " ").trim();
      return html.slice(0, 400);
    } catch (err) {
      return "";
    }
  }

  function pushEvent(type, node, extra) {
    if (!node || node.nodeType !== 1) return;
    var tag = (node.tagName || "").toLowerCase();
    var cls = node.className || "";
    var txt = shortText(node);
    var html = shortHtml(node);
    window.__SOMA_DEBUG_EVENTS.push({
      timestamp: now(),
      type: type,
      tag: tag,
      id: node.id || "",
      class: cls,
      text: txt,
      html: html,
      url: window.location.href,
      extra: extra || {}
    });
  }

  function interesting(node) {
    if (!node || node.nodeType !== 1) return false;
    var tag = (node.tagName || "").toLowerCase();
    var cls = (node.className || "").toString();
    if (cls.indexOf("swal") !== -1 || cls.indexOf("swal2") !== -1) return true;
    if (tag === "button") return true;
    if (tag === "input") {
      var type = (node.getAttribute("type") || "").toLowerCase();
      return type === "button" || type === "submit";
    }
    if (tag === "a" && (node.getAttribute("role") || "").toLowerCase() === "button") return true;
    return false;
  }

  function scanNode(node, type) {
    if (!interesting(node)) {
      var matches = false;
      try {
        matches = !!(node && node.querySelector && node.querySelector('[class*="swal"], [class*="swal2"], button, input[type="button"], input[type="submit"], a[role="button"]'));
      } catch (err) {
        matches = false;
      }
      if (!matches) return;
      pushEvent(type, node, { matched: "descendant" });
      return;
    }
    pushEvent(type, node);
  }

  var observer = new MutationObserver(function (mutations) {
    var currentUrl = window.location.href;
    if (currentUrl !== window.__SOMA_DEBUG_LAST_URL) {
      window.__SOMA_DEBUG_EVENTS.push({
        timestamp: now(),
        type: "url",
        tag: "location",
        id: "",
        class: "",
        text: document.title || "",
        html: "",
        url: currentUrl,
        extra: { previous: window.__SOMA_DEBUG_LAST_URL }
      });
      window.__SOMA_DEBUG_LAST_URL = currentUrl;
    }

    mutations.forEach(function (mutation) {
      if (mutation.type === "childList") {
        mutation.addedNodes.forEach(function (node) { scanNode(node, "added"); });
        mutation.removedNodes.forEach(function (node) { scanNode(node, "removed"); });
      } else if (mutation.type === "attributes") {
        scanNode(mutation.target, "attributes");
      } else if (mutation.type === "characterData") {
        scanNode(mutation.target.parentElement || mutation.target, "text");
      }
    });
  });

  observer.observe(document.body, {
    subtree: true,
    childList: true,
    attributes: true,
    characterData: true,
    attributeFilter: ["class", "style", "aria-hidden", "role", "id"],
  });

  var originalPushState = history.pushState;
  history.pushState = function () {
    var result = originalPushState.apply(this, arguments);
    window.__SOMA_DEBUG_EVENTS.push({
      timestamp: now(),
      type: "url",
      tag: "history.pushState",
      id: "",
      class: "",
      text: document.title || "",
      html: "",
      url: window.location.href,
      extra: {}
    });
    window.__SOMA_DEBUG_LAST_URL = window.location.href;
    return result;
  };

  var originalReplaceState = history.replaceState;
  history.replaceState = function () {
    var result = originalReplaceState.apply(this, arguments);
    window.__SOMA_DEBUG_EVENTS.push({
      timestamp: now(),
      type: "url",
      tag: "history.replaceState",
      id: "",
      class: "",
      text: document.title || "",
      html: "",
      url: window.location.href,
      extra: {}
    });
    window.__SOMA_DEBUG_LAST_URL = window.location.href;
    return result;
  };

  window.addEventListener("hashchange", function () {
    window.__SOMA_DEBUG_EVENTS.push({
      timestamp: now(),
      type: "url",
      tag: "hashchange",
      id: "",
      class: "",
      text: document.title || "",
      html: "",
      url: window.location.href,
      extra: {}
    });
    window.__SOMA_DEBUG_LAST_URL = window.location.href;
  });

  window.__SOMA_DEBUG_MUTATION_OBSERVER_INSTALLED = true;
  return "ok";
})();
"""
        try:
            self.driver.execute_script(script)
            self._observer_installed = True
            self._write_log("[DEBUG] MutationObserver instalado")
        except Exception as exc:
            self._write_log(f"[DEBUG][WARN] Falha ao instalar MutationObserver: {exc}")

    @staticmethod
    def _truncate(value: Any, max_len: int = 220) -> str:
        text = "" if value is None else str(value)
        text = " ".join(text.split())
        return text if len(text) <= max_len else text[:max_len] + "…"

    def _redact(self, value: Any) -> str:
        text = self._truncate(value, 1000)
        secret = self._secret.strip()
        if secret:
            text = text.replace(secret, "[REDACTED]")
        return text

    def _element_summary(self, element: Any) -> dict[str, Any]:
        summary: dict[str, Any] = {}
        attrs = ("tag_name", "id", "class", "name", "text", "value")
        for attr in attrs:
            try:
                if attr == "tag_name":
                    summary["tag"] = getattr(element, "tag_name", "") or ""
                elif attr == "text":
                    summary["text"] = self._redact(getattr(element, "text", ""))
                else:
                    summary[attr] = self._redact(element.get_attribute(attr))
            except Exception:
                summary[attr if attr != "tag_name" else "tag"] = ""
        try:
            summary["displayed"] = bool(element.is_displayed())
        except Exception:
            summary["displayed"] = False
        try:
            summary["enabled"] = bool(element.is_enabled())
        except Exception:
            summary["enabled"] = False
        try:
            summary["outerHTML"] = self._redact(element.get_attribute("outerHTML"))
        except Exception:
            summary["outerHTML"] = ""
        return summary

    def _print_element_list(self, title: str, elements: Iterable[Any]) -> str:
        items = list(elements)
        lines = [title, f"FOUND={len(items)}"]
        for idx, element in enumerate(items):
            summary = self._element_summary(element)
            lines.append(
                f"[{idx}] tag={summary.get('tag','')} displayed={summary.get('displayed')} enabled={summary.get('enabled')} "
                f"text={summary.get('text','')} id={summary.get('id','')} name={summary.get('name','')} "
                f"class={summary.get('class','')} value={summary.get('value','')}"
            )
            lines.append(f"outerHTML={summary.get('outerHTML','')}")
        output = "\n".join(lines)
        print(output)
        self._write_log(output)
        return output

    def _xpath_for_element(self, element: Any) -> str:
        script = r"""
function absoluteXPath(el) {
  if (!el || el.nodeType !== 1) return "";
  if (el === document.body) return "/html/body";
  if (el === document.documentElement) return "/html";
  var segs = [];
  while (el && el.nodeType === 1 && el !== document.documentElement) {
    var ix = 1;
    var sib = el.previousElementSibling;
    while (sib) {
      if (sib.tagName === el.tagName) ix += 1;
      sib = sib.previousElementSibling;
    }
    segs.unshift(el.tagName.toLowerCase() + "[" + ix + "]");
    if (el.parentElement === document.documentElement) {
      segs.unshift("html");
      break;
    }
    el = el.parentElement;
  }
  var path = "/" + segs.join("/");
  return path.replace("/html[1]", "/html");
}
return absoluteXPath(arguments[0]);
"""
        try:
            return str(self.driver.execute_script(script, element) or "")
        except Exception:
            return ""

    def _dump_result(self, title: str, result: Any) -> str:
        output = f"{title}\n{self._redact(result)}"
        print(output)
        self._write_log(output)
        return output

    def _cmd_xpath(self, xpath: str) -> str:
        elements = self.driver.find_elements(By.XPATH, xpath)
        return self._print_element_list(f"XPath: {xpath}", elements)

    def _cmd_css(self, selector: str) -> str:
        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
        return self._print_element_list(f"CSS: {selector}", elements)

    def _cmd_swal(self) -> str:
        elements = []
        try:
            elements.extend(self.driver.find_elements(By.CSS_SELECTOR, "[class*='swal'], [class*='swal2']"))
        except Exception:
            pass
        return self._print_element_list("SweetAlert elements", elements)

    def _cmd_buttons(self) -> str:
        elements: list[Any] = []
        selectors = [
            (By.TAG_NAME, "button"),
            (By.CSS_SELECTOR, "input[type='button']"),
            (By.CSS_SELECTOR, "input[type='submit']"),
            (By.CSS_SELECTOR, "a[role='button']"),
        ]
        for by, value in selectors:
            try:
                elements.extend(self.driver.find_elements(by, value))
            except Exception:
                continue
        visible = []
        for element in elements:
            try:
                if element.is_displayed():
                    visible.append(element)
            except Exception:
                continue
        lines = ["Buttons visíveis:"]
        lines.append(f"FOUND={len(visible)}")
        for idx, element in enumerate(visible):
            summary = self._element_summary(element)
            abs_xpath = self._xpath_for_element(element)
            lines.append(
                f"[{idx}] tag={summary.get('tag','')} text={summary.get('text','')} id={summary.get('id','')} "
                f"class={summary.get('class','')} xpath={abs_xpath}"
            )
            lines.append(f"outerHTML={summary.get('outerHTML','')}")
        output = "\n".join(lines)
        print(output)
        self._write_log(output)
        return output

    def _cmd_url(self) -> str:
        current_url = getattr(self.driver, "current_url", "")
        title = getattr(self.driver, "title", "")
        ready_state = ""
        try:
            ready_state = str(self.driver.execute_script("return document.readyState;") or "")
        except Exception:
            ready_state = ""
        output = "\n".join(
            [
                f"URL={current_url}",
                f"TITLE={title}",
                f"READY_STATE={ready_state}",
            ]
        )
        print(output)
        self._write_log(output)
        return output

    def _cmd_shot(self) -> str:
        if hasattr(self.actions, "screenshot"):
            path = self.actions.screenshot(f"soma_debug_{dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}")
            output = f"Screenshot={path}"
        else:
            output = "Screenshot indisponível"
        print(output)
        self._write_log(output)
        return output

    def _cmd_html(self) -> str:
        if hasattr(self.actions, "dump_page_source"):
            path = self.actions.dump_page_source(f"soma_debug_{dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}")
            output = f"HTML={path}"
        else:
            output = "HTML indisponível"
        print(output)
        self._write_log(output)
        return output

    def _cmd_frames(self) -> str:
        frames: list[Any] = []
        for by, value in ((By.TAG_NAME, "iframe"), (By.TAG_NAME, "frame")):
            try:
                frames.extend(self.driver.find_elements(by, value))
            except Exception:
                continue
        lines = ["Frames visíveis:"]
        lines.append(f"FOUND={len(frames)}")
        for idx, element in enumerate(frames):
            summary = self._element_summary(element)
            lines.append(
                f"[{idx}] tag={summary.get('tag','')} id={summary.get('id','')} name={summary.get('name','')} "
                f"class={summary.get('class','')} displayed={summary.get('displayed')}"
            )
            lines.append(f"src={self._truncate(element.get_attribute('src'), 240)}")
        output = "\n".join(lines)
        print(output)
        self._write_log(output)
        return output

    def _cmd_events(self) -> str:
        try:
            events = self.driver.execute_script("return window.__SOMA_DEBUG_EVENTS || [];") or []
        except Exception as exc:
            output = f"events indisponíveis: {exc}"
            print(output)
            self._write_log(output)
            return output

        lines = ["Eventos DOM/SweetAlert:"]
        lines.append(f"FOUND={len(events)}")
        for item in events[-100:]:
            if not isinstance(item, dict):
                lines.append(self._truncate(item, 500))
                continue
            lines.append(
                f"{item.get('timestamp', '')} | {item.get('type', '')} | {item.get('tag', '')} | "
                f"class={self._truncate(item.get('class', ''), 120)} | text={self._truncate(item.get('text', ''), 160)} | url={item.get('url', '')}"
            )
        output = "\n".join(lines)
        print(output)
        self._write_log(output)
        return output

    def _print_banner(
        self,
        *,
        stage: str,
        phase: str,
        action: str,
        element_name: str | None,
        locator: tuple[str, str] | None,
        value: Any,
        instructions: list[str],
    ) -> None:
        self._counter += 1
        step_id = f"{self._counter:02d}"
        lines = [
            "=" * 60,
            f"[DEBUG SOMA - PASSO {step_id}]",
            f"ETAPA: {stage}",
            f"MOMENTO: {'ANTES DA AÇÃO' if phase.upper() == 'BEFORE' else 'DEPOIS DA AÇÃO'}",
            f"AÇÃO: {action}",
        ]
        if element_name:
            lines.append(f"ELEMENTO: {element_name}")
        if locator:
            lines.extend([f"XPATH:\n{locator[1]}", f"METHOD: {locator[0]}"])
        if value not in (None, ""):
            lines.append(f"VALOR:\n{self._redact(value)}")
        if instructions:
            lines.append("O QUE FAZER AGORA:")
            for idx, item in enumerate(instructions, start=1):
                lines.append(f"{idx}. {self._redact(item)}")
        lines.extend(
            [
                "",
                "COMANDOS DISPONÍVEIS:",
                "ENTER = executar próxima ação",
                "x <xpath> = testar XPath na sessão atual",
                "css <selector> = testar CSS apenas para diagnóstico",
                "swal = inspecionar SweetAlert",
                "buttons = listar botões visíveis",
                "events = listar eventos do MutationObserver",
                "url = mostrar URL / título / readyState",
                "shot = tirar screenshot",
                "html = salvar DOM atual",
                "frames = listar iframes",
                "q = interromper debug",
                "DEBUG>",
                "=" * 60,
            ]
        )
        banner = "\n".join(lines)
        print(banner)
        self._write_log(banner)

    def _handle_command(self, command: str) -> bool:
        cmd = command.strip()
        if not cmd:
            self._write_log("[DEBUG] ENTER")
            return True

        lowered = cmd.lower()
        self._write_log(f"[DEBUG] COMANDO={cmd}")

        if lowered == "q":
            self._quit_requested = True
            self._dump_result("DEBUG encerrado pelo utilizador.", "Browser permanece aberto até o fim normal do processo.")
            raise GuidedDebugAbort("DEBUG_STEP_MODE interrompido pelo utilizador.")
        if lowered.startswith("x "):
            self._cmd_xpath(cmd[2:].strip())
            return False
        if lowered.startswith("css "):
            self._cmd_css(cmd[4:].strip())
            return False
        if lowered == "swal":
            self._cmd_swal()
            return False
        if lowered == "buttons":
            self._cmd_buttons()
            return False
        if lowered == "events":
            self._cmd_events()
            return False
        if lowered == "url":
            self._cmd_url()
            return False
        if lowered == "shot":
            self._cmd_shot()
            return False
        if lowered == "html":
            self._cmd_html()
            return False
        if lowered == "frames":
            self._cmd_frames()
            return False

        self._dump_result(
            "Comando desconhecido.",
            "Use ENTER, x <xpath>, css <selector>, swal, buttons, events, url, shot, html, frames ou q.",
        )
        return False

    def checkpoint(
        self,
        *,
        row: Any,
        stage: str,
        phase: str,
        action: str,
        element_name: str | None = None,
        locator: tuple[str, str] | None = None,
        value: Any = None,
        instructions: Iterable[str] | None = None,
    ) -> None:
        row_number = getattr(row, "row_number", None)
        if not self.enabled_for_row(row_number):
            return

        self._ensure_observer()
        self._print_banner(
            stage=stage,
            phase=phase,
            action=action,
            element_name=element_name,
            locator=locator,
            value=value,
            instructions=list(instructions or ()),
        )

        while True:
            try:
                command = input("DEBUG> ")
            except EOFError:
                command = "q"

            if self._handle_command(command):
                return

    def checkpoint_before(
        self,
        *,
        row: Any,
        stage: str,
        action: str,
        element_name: str | None = None,
        locator: tuple[str, str] | None = None,
        value: Any = None,
        instructions: Iterable[str] | None = None,
    ) -> None:
        self.checkpoint(
            row=row,
            stage=stage,
            phase="BEFORE",
            action=action,
            element_name=element_name,
            locator=locator,
            value=value,
            instructions=instructions,
        )

    def checkpoint_after(
        self,
        *,
        row: Any,
        stage: str,
        action: str,
        element_name: str | None = None,
        locator: tuple[str, str] | None = None,
        value: Any = None,
        instructions: Iterable[str] | None = None,
    ) -> None:
        self.checkpoint(
            row=row,
            stage=stage,
            phase="AFTER",
            action=action,
            element_name=element_name,
            locator=locator,
            value=value,
            instructions=instructions,
        )
