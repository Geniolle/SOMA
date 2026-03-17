from __future__ import annotations

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
        return WebDriverWait(self.driver, timeout_seconds or self.cfg.timeout_seconds)

    def screenshot(self, name: str) -> Path:
        safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in name)
        path = self.cfg.screenshots_dir / f"{safe}.png"
        try:
            self.driver.save_screenshot(str(path))
        except (InvalidSessionIdException, WebDriverException) as e:
            log.error("Screenshot falhou (sessão do browser não está ativa): %s", e)
        return path

    def wait_dom_ready(self, timeout_seconds: int = 30) -> None:
        end = time.time() + timeout_seconds
        while time.time() < end:
            try:
                state = self.driver.execute_script("return document.readyState")
                if state == "complete":
                    return
            except Exception:
                pass
            time.sleep(0.2)
        log.warning("DOM não ficou 'complete' dentro de %ss", timeout_seconds)

    def wait_present(self, locator: Locator, timeout_seconds: Optional[int] = None) -> WebElement:
        return self._wait(timeout_seconds).until(EC.presence_of_element_located(locator))

    def wait_visible(self, locator: Locator, timeout_seconds: Optional[int] = None) -> WebElement:
        return self._wait(timeout_seconds).until(EC.visibility_of_element_located(locator))

    def wait_clickable(self, locator: Locator, timeout_seconds: Optional[int] = None) -> WebElement:
        return self._wait(timeout_seconds).until(EC.element_to_be_clickable(locator))

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
        try:
            self._wait(timeout_seconds).until(EC.presence_of_element_located(locator))
            return True
        except TimeoutException:
            return False

    def click(self, locator: Locator) -> None:
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] click | by=%s sel=%s", locator[0], locator[1])
        el = self.wait_clickable(locator)
        try:
            el.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", el)

    def click_js(self, locator: Locator) -> None:
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] click_js | by=%s sel=%s", locator[0], locator[1])
        el = self.wait_present(locator, timeout_seconds=30)
        try:
            self.driver.execute_script("arguments[0].click();", el)
        except Exception:
            el.click()

    def type(self, locator: Locator, text: str, clear: bool = True) -> None:
        if log.isEnabledFor(logging.DEBUG):
            log.debug(
                "[ACTION] type | by=%s sel=%s | clear=%s | len=%s",
                locator[0],
                locator[1],
                clear,
                len(str(text or "")),
            )
        el = self.wait_visible(locator)
        if clear:
            el.clear()
        el.send_keys(text)

    def press_enter(self, locator: Locator) -> None:
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] press_enter | by=%s sel=%s", locator[0], locator[1])
        el = self.wait_visible(locator)
        el.send_keys(Keys.ENTER)

    def select_by_text(self, locator: Locator, text: str) -> None:
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] select_by_text | by=%s sel=%s | text=%s", locator[0], locator[1], str(text)[:60])
        el = self.wait_visible(locator)
        Select(el).select_by_visible_text(text)

    def wait_invisible(self, locator: Locator, timeout_seconds: Optional[int] = None) -> None:
        self._wait(timeout_seconds).until(EC.invisibility_of_element_located(locator))

    def select2_choose(
        self,
        opener: Locator,
        value: str,
        search_input: Locator = (By.XPATH, "/html/body/span/span/span[1]/input"),
    ) -> None:
        """
        Select2 robusto:
        - Primeiro tenta input de pesquisa.
        - Se não existir pesquisa, seleciona clicando numa opção da lista.
        """
        if log.isEnabledFor(logging.DEBUG):
            log.debug("[ACTION] select2_choose | opener=%s | value=%s", opener, str(value)[:60])

        # abrir dropdown
        try:
            self.click(opener)
        except Exception:
            self.click_js(opener)

        css_search = (By.CSS_SELECTOR, "span.select2-container--open input.select2-search__field")
        css_options = (By.CSS_SELECTOR, "li.select2-results__option")

        # tenta pesquisa
        for loc in (css_search, search_input):
            if self.exists(loc, timeout_seconds=2):
                inp = self.wait_visible(loc, timeout_seconds=10)
                inp.clear()
                inp.send_keys(value)
                inp.send_keys(Keys.ENTER)
                return

        # sem pesquisa: clicar opção
        try:
            self._wait(10).until(lambda d: len(d.find_elements(*css_options)) > 0)
        except TimeoutException:
            p = self.screenshot("select2_no_options")
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
            raise RuntimeError(f"Opção Select2 não encontrada: '{value}'. Amostra={sample} | screenshot={p}")

        try:
            self.driver.execute_script("arguments[0].click();", best)
        except Exception:
            best.click()