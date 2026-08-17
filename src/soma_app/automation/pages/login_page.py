from __future__ import annotations

import logging
import time

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from soma_app.automation.actions import Actions
from soma_app.config.locators import apply_locator_overrides
from soma_app.config.settings import Settings
from soma_app.infra.env import env_bool
from soma_app.infra.trace import log_kv, step

log = logging.getLogger("soma_app.pages.login")


class LoginPage:
    EMAIL = (By.NAME, "email")
    SENHA = (By.NAME, "senha")
    SUBMIT = (By.NAME, "submit")

    SOMA_READY = (By.XPATH, "")

    SOMA_BUTTON_CANDIDATES = []

    def __init__(self, actions: Actions, settings: Settings):
        self.a = actions
        self.settings = settings
        apply_locator_overrides(self, "login")

    def login(self) -> None:
        with step(log, "login.open_portal", url=self.settings.site_login_url):
            self.a.driver.get(self.settings.site_login_url)

        self._fill_credentials(debug=env_bool("DEBUG_STEP_MODE", default=False))

        with step(log, "login.wait_form_disappear"):
            try:
                self.a.wait_any_present(self.SOMA_BUTTON_CANDIDATES, timeout_seconds=10)
            except Exception:
                log.debug(
                    "Portal ainda a estabilizar apos submit; seguindo para a etapa de abertura do SOMA.",
                    extra={"url": self.a.driver.current_url, "title": self.a.driver.title},
                )

        with step(log, "login.open_soma_app"):
            self.open_soma_app()

        log.info("Login OK e SOMA pronto | url=%s | title=%s", self.a.driver.current_url, self.a.driver.title)

    def ensure_soma_home(self) -> None:
        with step(log, "soma.ensure_home", url=self.a.driver.current_url, title=self.a.driver.title):
            if self.a.exists(self.SOMA_READY, timeout_seconds=2):
                return
            self.open_soma_app()

    def _fill_credentials(self, *, debug: bool) -> None:
        with step(log, "login.fill_credentials"):
            if not debug:
                self.a.type(self.EMAIL, self.settings.site_user)
                self.a.type(self.SENHA, self.settings.site_password)
                self.a.click(self.SUBMIT)
                return

            self._type_login_field(self.EMAIL, self.settings.site_user)
            self._type_login_field(self.SENHA, self.settings.site_password)
            self.a.click(self.SUBMIT)

    def _type_login_field(self, locator: tuple[str, str], value: str) -> None:
        try:
            self.a.type(locator, value)
            return
        except TimeoutException:
            el = self.a.wait_present(locator, timeout_seconds=max(30, self.settings.timeout_seconds))
            try:
                self.a.driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center', inline:'center'});",
                    el,
                )
            except Exception:
                pass

            try:
                el.click()
                el.clear()
                el.send_keys(value)
                return
            except Exception:
                self.a.driver.execute_script(
                    """
                    const el = arguments[0];
                    const value = arguments[1];
                    el.focus();
                    el.value = value;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    """,
                    el,
                    value,
                )

    def open_soma_app(self) -> None:
        with step(log, "soma.portal_open", url=self.settings.site_login_url):
            try:
                self.a.driver.get(self.settings.site_login_url)
            except Exception:
                pass

        before_handles = list(self.a.driver.window_handles)

        with step(log, "soma.find_button"):
            try:
                winner = self.a.wait_any_present(
                    self.SOMA_BUTTON_CANDIDATES,
                    timeout_seconds=max(30, self.settings.timeout_seconds),
                )
            except Exception:
                p = self.a.screenshot("soma_button_not_found")
                log_kv(
                    log,
                    "Botao SOMA nao encontrado.",
                    level=logging.ERROR,
                    url=self.a.driver.current_url,
                    title=self.a.driver.title,
                    screenshot=p,
                )
                raise

        with step(log, "soma.click_button", locator=str(winner)):
            self.a.click_js(winner)

        time.sleep(1)
        after_handles = list(self.a.driver.window_handles)
        if len(after_handles) > len(before_handles):
            new_handle = [h for h in after_handles if h not in before_handles][-1]
            with step(log, "soma.switch_window", new_handle=new_handle):
                self.a.driver.switch_to.window(new_handle)

        with step(log, "soma.wait_ready"):
            try:
                self.a.wait_visible(self.SOMA_READY, timeout_seconds=max(60, self.settings.timeout_seconds))
            except TimeoutException:
                try:
                    self.a.wait_present(self.SOMA_READY, timeout_seconds=10)
                    log.warning(
                        "SOMA_READY nao ficou visivel, mas o seletor foi encontrado no DOM; seguindo.",
                        extra={"url": self.a.driver.current_url, "title": self.a.driver.title},
                    )
                except TimeoutException:
                    p = self.a.screenshot("soma_not_ready")
                    log_kv(
                        log,
                        "SOMA nao carregou a tempo.",
                        level=logging.ERROR,
                        url=self.a.driver.current_url,
                        title=self.a.driver.title,
                        screenshot=p,
                    )
                    raise RuntimeError("SOMA nao carregou (Entradas/saidas nao apareceu).")
