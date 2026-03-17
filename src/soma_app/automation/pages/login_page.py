from __future__ import annotations

import logging
import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

from soma_app.automation.actions import Actions
from soma_app.config.settings import Settings
from soma_app.infra.trace import step, log_kv

log = logging.getLogger("soma_app.pages.login")


class LoginPage:
    EMAIL = (By.NAME, "email")
    SENHA = (By.NAME, "senha")
    SUBMIT = (By.NAME, "submit")

    SOMA_READY = (By.XPATH, '//span[contains(text(),"Entradas/saídas")]')

    SOMA_BUTTON_CANDIDATES = [
        (By.ID, "285"),
        (By.XPATH, "//*[@id='285']"),
        (By.XPATH, "//*[self::a or self::button][contains(.,'SOMA') or contains(.,'Soma')]"),
    ]

    def __init__(self, actions: Actions, settings: Settings):
        self.a = actions
        self.settings = settings

    def login(self) -> None:
        with step(log, "login.open_portal", url=self.settings.site_login_url):
            self.a.driver.get(self.settings.site_login_url)

        with step(log, "login.fill_credentials"):
            self.a.type(self.EMAIL, self.settings.site_user)
            self.a.type(self.SENHA, self.settings.site_password)
            self.a.click(self.SUBMIT)

        with step(log, "login.wait_form_disappear"):
            try:
                self.a.wait_invisible(self.EMAIL, timeout_seconds=45)
            except Exception:
                p = self.a.screenshot("login_submit_no_redirect")
                log_kv(
                    log,
                    logging.ERROR,
                    "Login não avançou (form ainda visível).",
                    url=self.a.driver.current_url,
                    title=self.a.driver.title,
                    screenshot=p,
                )
                raise RuntimeError("Login não avançou (form ainda visível).")

        with step(log, "login.open_soma_app"):
            self.open_soma_app()

        log.info("Login OK e SOMA pronto | url=%s | title=%s", self.a.driver.current_url, self.a.driver.title)

    def ensure_soma_home(self) -> None:
        with step(log, "soma.ensure_home", url=self.a.driver.current_url, title=self.a.driver.title):
            if self.a.exists(self.SOMA_READY, timeout_seconds=2):
                return
            self.open_soma_app()

    def open_soma_app(self) -> None:
        # garante estar no portal (quando necessário)
        with step(log, "soma.portal_open", url=self.settings.site_login_url):
            try:
                self.a.driver.get(self.settings.site_login_url)
            except Exception:
                pass

        before_handles = list(self.a.driver.window_handles)

        with step(log, "soma.find_button"):
            try:
                winner = self.a.wait_any_present(self.SOMA_BUTTON_CANDIDATES, timeout_seconds=max(30, self.settings.timeout_seconds))
            except Exception:
                p = self.a.screenshot("soma_button_not_found")
                log_kv(log, logging.ERROR, "Botão SOMA não encontrado.", url=self.a.driver.current_url, title=self.a.driver.title, screenshot=p)
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
                p = self.a.screenshot("soma_not_ready")
                log_kv(log, logging.ERROR, "SOMA não carregou a tempo.", url=self.a.driver.current_url, title=self.a.driver.title, screenshot=p)
                raise RuntimeError("SOMA não carregou (Entradas/saídas não apareceu).")