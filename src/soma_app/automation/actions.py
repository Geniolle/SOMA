from __future__ import annotations

import json
import logging
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple

from selenium.common.exceptions import (
    ElementClickInterceptedException,
    InvalidSessionIdException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from soma_app.infra.audit import audit_event

log = logging.getLogger("soma_app.automation.actions")

Locator = Tuple[str, str]


@dataclass(frozen=True)
class ActionConfig:
    timeout_seconds: int = 20
    screenshots_dir: Path = Path("artifacts/screenshots")


class Actions:
    def __init__(self, driver: WebDriver, cfg: ActionConfig):
        self.driver = driver
        self.cfg = cfg
        self.cfg.screenshots_dir.mkdir(parents=True, exist_ok=True)

    def _wait(self, timeout_seconds: Optional[int] = None) -> WebDriverWait:
        if timeout_seconds is None:
            timeout_seconds = self.cfg.timeout_seconds
        return WebDriverWait(self.driver, timeout_seconds)

    @staticmethod
    def _loc(locator: Locator) -> str:
        return f"{locator[0]}::{locator[1]}"

    @staticmethod
    def _preview(value: object, max_len: int = 80) -> str:
        s = "" if value is None else str(value)
        s = " ".join(s.split())
        return s[:max_len]

    def _audit(self, event: str, **fields: object) -> None:
        try:
            audit_event(event, **fields)
        except Exception:
            pass

    def screenshot(self, name: str) -> Path:
        safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in name)
        path = self.cfg.screenshots_dir / f"{safe}.png"
        try:
            self.driver.save_screenshot(str(path))
        except (InvalidSessionIdException, WebDriverException) as e:
            log.error("Screenshot falhou (sessão do browser não está ativa): %s", e)
        return path

    def dump_page_source(self, name: str) -> Path:
        safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in name)
        path = self.cfg.screenshots_dir.parent / "diagnostics" / f"{safe}.html"
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            path.write_text(self.driver.page_source or "", encoding="utf-8")
        except Exception as e:
            log.error("Falha ao gravar page_source: %s", e)
        return path

    def dump_locator_probe(self, name: str, locators: Iterable[Locator]) -> Path:
        safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in name)
        path = self.cfg.screenshots_dir.parent / "diagnostics" / f"{safe}.json"
        path.parent.mkdir(parents=True, exist_ok=True)

        payload = []
        for locator in locators:
            entry: dict[str, object] = {"locator": locator, "count": 0, "found": False}
            try:
                elements = self.driver.find_elements(*locator)
                entry["count"] = len(elements)
                if elements:
                    el = elements[0]
                    entry["found"] = True
                    try:
                        entry["tag"] = el.tag_name
                    except Exception:
                        pass
                    try:
                        entry["text"] = (el.text or "").strip()
                    except Exception:
                        pass
                    try:
                        entry["value"] = (el.get_attribute("value") or "").strip()
                    except Exception:
                        pass
                    try:
                        entry["name"] = el.get_attribute("name")
                    except Exception:
                        pass
                    try:
                        entry["id"] = el.get_attribute("id")
                    except Exception:
                        pass
                    try:
                        entry["class"] = el.get_attribute("class")
                    except Exception:
                        pass
                    try:
                        entry["displayed"] = bool(el.is_displayed())
                    except Exception:
                        pass
                    try:
                        entry["enabled"] = bool(el.is_enabled())
                    except Exception:
                        pass
                    try:
                        entry["outer_html"] = self.driver.execute_script("return arguments[0].outerHTML;", el)
                    except Exception:
                        pass
            except Exception as e:
                entry["error"] = str(e)
            payload.append(entry)

        try:
            path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            log.error("Falha ao gravar probe de locators: %s", e)
        return path

    def wait_dom_ready(self, timeout_seconds: int = 30) -> None:
        t0 = time.perf_counter()
        self._audit("UI_WAIT_DOM_READY_START", timeout_seconds=timeout_seconds)
        end = time.time() + timeout_seconds
        while time.time() < end:
            try:
                state = self.driver.execute_script("return document.readyState")
                if state == "complete":
                    self._audit("UI_WAIT_DOM_READY_OK", timeout_seconds=timeout_seconds, dt_ms=int((time.perf_counter() - t0) * 1000))
                    return
            except Exception:
                pass
            time.sleep(0.2)
        log.warning("DOM não ficou 'complete' dentro de %ss", timeout_seconds)
        self._audit("UI_WAIT_DOM_READY_TIMEOUT", timeout_seconds=timeout_seconds, dt_ms=int((time.perf_counter() - t0) * 1000))

    def wait_present(self, locator: Locator, timeout_seconds: Optional[int] = None) -> WebElement:
        t0 = time.perf_counter()
        timeout = timeout_seconds if timeout_seconds is not None else self.cfg.timeout_seconds
        self._audit("UI_WAIT_PRESENT_START", locator=self._loc(locator), timeout_seconds=timeout)
        try:
            el = self._wait(timeout_seconds).until(EC.presence_of_element_located(locator))
            self._audit("UI_WAIT_PRESENT_OK", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000))
            return el
        except Exception as e:
            self._audit("UI_WAIT_PRESENT_FAIL", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000), err=type(e).__name__)
            raise

    def wait_visible(self, locator: Locator, timeout_seconds: Optional[int] = None) -> WebElement:
        t0 = time.perf_counter()
        timeout = timeout_seconds if timeout_seconds is not None else self.cfg.timeout_seconds
        self._audit("UI_WAIT_VISIBLE_START", locator=self._loc(locator), timeout_seconds=timeout)
        try:
            el = self._wait(timeout_seconds).until(EC.visibility_of_element_located(locator))
            self._audit("UI_WAIT_VISIBLE_OK", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000))
            return el
        except Exception as e:
            self._audit("UI_WAIT_VISIBLE_FAIL", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000), err=type(e).__name__)
            raise

    def wait_clickable(self, locator: Locator, timeout_seconds: Optional[int] = None) -> WebElement:
        t0 = time.perf_counter()
        timeout = timeout_seconds if timeout_seconds is not None else self.cfg.timeout_seconds
        self._audit("UI_WAIT_CLICKABLE_START", locator=self._loc(locator), timeout_seconds=timeout)
        try:
            el = self._wait(timeout_seconds).until(EC.element_to_be_clickable(locator))
            self._audit("UI_WAIT_CLICKABLE_OK", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000))
            return el
        except Exception as e:
            self._audit("UI_WAIT_CLICKABLE_FAIL", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000), err=type(e).__name__)
            raise

    def wait_any_present(self, locators: Iterable[Locator], timeout_seconds: Optional[int] = None) -> Locator:
        last_err: Optional[Exception] = None

        def _probe(_driver):
            nonlocal last_err
            for loc in locators:
                try:
                    _driver.find_element(*loc)
                    return loc
                except Exception as e:
                    last_err = e
            return False

        try:
            return self._wait(timeout_seconds).until(_probe)
        except TimeoutException:
            raise TimeoutException(f"Timeout à espera de qualquer locator presente. last_err={last_err}")

    def exists(self, locator: Locator, timeout_seconds: Optional[int] = None) -> bool:
        timeout = timeout_seconds if timeout_seconds is not None else self.cfg.timeout_seconds
        t0 = time.perf_counter()
        try:
            self._wait(timeout_seconds).until(EC.presence_of_element_located(locator))
            self._audit("UI_EXISTS_TRUE", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000))
            return True
        except TimeoutException:
            self._audit("UI_EXISTS_FALSE", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000))
            return False

    def click(self, locator: Locator) -> None:
        t0 = time.perf_counter()
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] click | by=%s sel=%s", locator[0], locator[1])
        self._audit("UI_CLICK_START", locator=self._loc(locator))
        el = self.wait_clickable(locator)
        try:
            el.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", el)
        self._audit("UI_CLICK_OK", locator=self._loc(locator), dt_ms=int((time.perf_counter() - t0) * 1000))

    def click_js(self, locator: Locator) -> None:
        t0 = time.perf_counter()
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] click_js | by=%s sel=%s", locator[0], locator[1])
        self._audit("UI_CLICK_JS_START", locator=self._loc(locator))
        el = self.wait_present(locator, timeout_seconds=30)
        try:
            self.driver.execute_script("arguments[0].click();", el)
        except Exception:
            el.click()
        self._audit("UI_CLICK_JS_OK", locator=self._loc(locator), dt_ms=int((time.perf_counter() - t0) * 1000))

    def type(self, locator: Locator, text: str, clear: bool = True) -> None:
        t0 = time.perf_counter()
        if log.isEnabledFor(logging.DEBUG):
            log.debug(
                "[ACTION] type | by=%s sel=%s | clear=%s | len=%s",
                locator[0],
                locator[1],
                clear,
                len(str(text or "")),
            )
        preview = self._preview(text)
        self._audit("UI_TYPE_START", locator=self._loc(locator), clear=clear, value_preview=preview, value_len=len(str(text or "")))
        el = self.wait_visible(locator)
        before = ""
        try:
            before = (el.get_attribute("value") or "").strip()
        except Exception:
            pass
        if clear:
            try:
                el.click()
            except Exception:
                pass
            try:
                el.send_keys(Keys.CONTROL, "a")
                el.send_keys(Keys.BACKSPACE)
            except Exception:
                try:
                    el.clear()
                except Exception:
                    pass
        el.send_keys(text)
        after = ""
        try:
            after = (el.get_attribute("value") or "").strip()
        except Exception:
            pass
        self._audit(
            "UI_TYPE_OK",
            locator=self._loc(locator),
            clear=clear,
            value_preview=preview,
            before_preview=self._preview(before),
            after_preview=self._preview(after),
            dt_ms=int((time.perf_counter() - t0) * 1000),
        )

    def press_enter(self, locator: Locator) -> None:
        t0 = time.perf_counter()
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] press_enter | by=%s sel=%s", locator[0], locator[1])
        self._audit("UI_PRESS_ENTER_START", locator=self._loc(locator))
        el = self.wait_visible(locator)
        el.send_keys(Keys.ENTER)
        self._audit("UI_PRESS_ENTER_OK", locator=self._loc(locator), dt_ms=int((time.perf_counter() - t0) * 1000))

    def select_by_text(self, locator: Locator, text: str) -> None:
        t0 = time.perf_counter()
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] select_by_text | by=%s sel=%s | text=%s", locator[0], locator[1], str(text)[:60])
        self._audit("UI_SELECT_TEXT_START", locator=self._loc(locator), text_preview=self._preview(text))
        el = self.wait_visible(locator)
        Select(el).select_by_visible_text(text)
        self._audit("UI_SELECT_TEXT_OK", locator=self._loc(locator), text_preview=self._preview(text), dt_ms=int((time.perf_counter() - t0) * 1000))

    def wait_invisible(self, locator: Locator, timeout_seconds: Optional[int] = None) -> None:
        t0 = time.perf_counter()
        timeout = timeout_seconds if timeout_seconds is not None else self.cfg.timeout_seconds
        self._audit("UI_WAIT_INVISIBLE_START", locator=self._loc(locator), timeout_seconds=timeout)
        try:
            self._wait(timeout_seconds).until(EC.invisibility_of_element_located(locator))
            self._audit("UI_WAIT_INVISIBLE_OK", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000))
        except Exception as e:
            self._audit("UI_WAIT_INVISIBLE_FAIL", locator=self._loc(locator), timeout_seconds=timeout, dt_ms=int((time.perf_counter() - t0) * 1000), err=type(e).__name__)
            raise

    def select2_choose(
        self,
        opener: Locator,
        value: str,
        search_input: Optional[Locator] = None,
    ) -> None:
        if search_input is None:
            from soma_app.config.locators import _coerce_locator, load_page_locator_config
            common_cfg = load_page_locator_config(None, "common")
            raw_search = common_cfg.get("SELECT2_SEARCH")
            if not raw_search:
                raise ValueError("O locator 'SELECT2_SEARCH' não está configurado na secção 'common' do JSON.")
            search_input = _coerce_locator(raw_search, None)
            if not search_input:
                raise ValueError("O locator 'SELECT2_SEARCH' na secção 'common' do JSON é inválido.")
        """
        Select2 robusto:
        - Primeiro tenta input de pesquisa.
        - Se não existir pesquisa, seleciona clicando numa opção da lista.
        """
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] select2_choose | opener=%s | value=%s", opener, str(value)[:60])
        self._audit("UI_SELECT2_START", opener=self._loc(opener), value_preview=self._preview(value))

        # abrir dropdown
        try:
            self.click(opener)
        except Exception as first_err:
            try:
                self.click_js(opener)
            except Exception as second_err:
                p = self.screenshot("select2_open_fail")
                self._audit(
                    "UI_SELECT2_FAIL",
                    opener=self._loc(opener),
                    value_preview=self._preview(value),
                    screenshot=str(p),
                    reason="open_fail",
                )
                raise TimeoutException(
                    f"Falha ao abrir Select2 | opener={opener} | value='{value}' | screenshot={p} | "
                    f"click_err={first_err} | click_js_err={second_err}"
                ) from second_err

        css_search = (By.CSS_SELECTOR, "span.select2-container--open input.select2-search__field")
        css_options = (By.CSS_SELECTOR, "li.select2-results__option")

        # tenta pesquisa
        for loc in (css_search, search_input):
            if self.exists(loc, timeout_seconds=2):
                inp = self.wait_visible(loc, timeout_seconds=10)
                inp.clear()
                inp.send_keys(value)
                inp.send_keys(Keys.ENTER)
                self._audit("UI_SELECT2_OK", opener=self._loc(opener), value_preview=self._preview(value), mode="search_input")
                return

        # sem pesquisa: clicar opção
        try:
            self._wait(10).until(lambda d: len(d.find_elements(*css_options)) > 0)
        except TimeoutException:
            p = self.screenshot("select2_no_options")
            self._audit(
                "UI_SELECT2_FAIL",
                opener=self._loc(opener),
                value_preview=self._preview(value),
                screenshot=str(p),
                reason="no_options",
            )
            raise TimeoutException(f"Select2 abriu sem pesquisa e sem opções visíveis | opener={opener} | screenshot={p}")

        options = self.driver.find_elements(*css_options)
        want = (value or "").strip().lower()

        best: Optional[WebElement] = None
        for opt in options:
            txt = (opt.text or "").strip()
            if not txt:
                continue
            tnorm = txt.lower()
            if tnorm == want:
                best = opt
                break
            if want and want in tnorm:
                best = opt

        if best is None:
            sample = [((o.text or "").strip()) for o in options[:10]]
            p = self.screenshot("select2_option_not_found")
            self._audit(
                "UI_SELECT2_FAIL",
                opener=self._loc(opener),
                value_preview=self._preview(value),
                screenshot=str(p),
                reason="option_not_found",
            )
            raise RuntimeError(f"Opção Select2 não encontrada: '{value}'. Amostra={sample} | screenshot={p}")

        try:
            self.driver.execute_script("arguments[0].click();", best)
        except Exception:
            best.click()
        self._audit("UI_SELECT2_OK", opener=self._loc(opener), value_preview=self._preview(value), mode="option_click")
