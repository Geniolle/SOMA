from __future__ import annotations

import logging
import time
from typing import Any, Dict, List, Optional, Tuple

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from soma_app.automation.actions import Actions
from soma_app.config.locators import apply_locator_overrides
from soma_app.infra.trace import log_kv, step

log = logging.getLogger("soma_app.pages.caixas_bancos")

Locator = Tuple[str, str]


class CaixasBancosPage:
    """
    Processo (6) portado do SYS_CAIXA.py, com robustez extra:
      - garante que a página Caixas/Bancos abriu (espera algum card/valor)
      - tenta reset (voltar ao portal + clicar SOMA) se necessário
      - lê valores com best-effort (não falha no primeiro card),
        mas falha se não conseguir ler NENHUM valor.
    """

    # --------- Portal / SOMA (para reset quando a UI estiver em estado estranho)
    SOMA_BUTTON_CANDIDATES: List[Locator] = []

    # --------- Menu Caixas/Bancos (xpath do SYS_CAIXA.py em primeiro)
    MENU_CAIXAS_BANCOS_CANDIDATES: List[Locator] = []

    # --------- Valores (xpaths do SYS_CAIXA.py)
    V_CAIXA_DIARIO = (By.XPATH, "")
    V_CAIXA_BANCO = (By.XPATH, "")
    V_CAIXA_CRIANCAS = (By.XPATH, "")
    V_CAIXA_CAFE = (By.XPATH, "")
    V_CAIXA_LIVRARIA = (By.XPATH, "")

    # Candidatos "qualquer valor apareceu?"
    ANY_VALUE_CANDIDATES: List[Locator] = []

    def __init__(self, actions: Actions, settings: Any):
        self.a = actions
        self.settings = settings
        self.timeout = int(getattr(settings, "timeout_seconds", 20) or 20)
        self.site_login_url = (getattr(settings, "site_login_url", "") or "https://verbodavida.info/apps/index.php").strip()
        apply_locator_overrides(self, "caixas_bancos")
        self.ANY_VALUE_CANDIDATES = [
            self.V_CAIXA_DIARIO,
            self.V_CAIXA_BANCO,
            self.V_CAIXA_CRIANCAS,
            self.V_CAIXA_CAFE,
            self.V_CAIXA_LIVRARIA,
        ]

    # -------------------------
    # helpers
    # -------------------------
    def _snap(self, name: str) -> str:
        try:
            p = self.a.screenshot(name)
            return str(p)
        except Exception:
            return ""

    def _reset_to_soma(self) -> None:
        """
        Volta ao portal e clica SOMA novamente.
        Isto resolve quando o browser ficou num "form/list" sem o menu lateral funcional.
        """
        with step(log, "caixas.reset_to_soma", url=self.site_login_url):
            try:
                self.a.driver.get(self.site_login_url)
                self.a.wait_dom_ready(15)
                time.sleep(1)
            except Exception:
                return

            try:
                loc = self.a.wait_any_present(self.SOMA_BUTTON_CANDIDATES, timeout_seconds=max(30, self.timeout))
                self.a.click_js(loc)
                self.a.wait_dom_ready(15)
                time.sleep(2)
            except Exception:
                p = self._snap("caixas_reset_fail")
                log_kv(log, "Falha ao reabrir SOMA.", level=logging.ERROR, url=self.a.driver.current_url, title=self.a.driver.title, screenshot=p)

    def _open_menu(self) -> None:
        with step(log, "caixas.open_menu"):
            loc = self.a.wait_any_present(self.MENU_CAIXAS_BANCOS_CANDIDATES, timeout_seconds=max(30, self.timeout))
            try:
                self.a.click_js(loc)
            except Exception:
                self.a.click(loc)
            self.a.wait_dom_ready(15)
            time.sleep(2)

    def _wait_any_value(self) -> None:
        # espera algum card/valor aparecer (sinal de que a página abriu)
        self.a.wait_any_present(self.ANY_VALUE_CANDIDATES, timeout_seconds=max(30, self.timeout))

    def open(self) -> None:
        print("\n(6.1) Abrindo página Caixas/Bancos...")

        # 1) tentativa normal
        try:
            self._open_menu()
            self._wait_any_value()
            return
        except Exception:
            p = self._snap("caixas_open_first_try_fail")
            log_kv(
                log,
                "Caixas/Bancos não abriu na 1ª tentativa. Vou tentar reset.",
                level=logging.ERROR,
                url=self.a.driver.current_url,
                title=self.a.driver.title,
                screenshot=p,
            )

        # 2) reset e tentar de novo
        self._reset_to_soma()
        self._open_menu()
        self._wait_any_value()

    def _read_value(self, label: str, locator: Locator) -> Optional[str]:
        try:
            # ajuda quando o card está fora do viewport
            try:
                self.a.scroll_into_view(locator)
            except Exception:
                pass

            el = self.a.wait_visible(locator, timeout_seconds=max(30, self.timeout))
            txt = (el.text or "").strip()
            if not txt:
                return None
            return txt
        except TimeoutException:
            p = self._snap(f"caixas_read_timeout_{label.lower().replace(' ', '_')}")
            log_kv(
                log,
                "Timeout ao ler valor.",
                level=logging.ERROR,
                label=label,
                url=self.a.driver.current_url,
                title=self.a.driver.title,
                screenshot=p,
            )
            return None
        except Exception as e:
            p = self._snap(f"caixas_read_fail_{label.lower().replace(' ', '_')}")
            log_kv(
                log,
                "Erro ao ler valor.",
                level=logging.ERROR,
                label=label,
                err=str(e)[:200],
                url=self.a.driver.current_url,
                title=self.a.driver.title,
                screenshot=p,
            )
            return None

    def read_values(self) -> Dict[str, str]:
        with step(log, "caixas.read_values"):
            vals: Dict[str, str] = {}

            v = self._read_value("CAIXA DIÁRIO", self.V_CAIXA_DIARIO)
            if v is not None:
                vals["CAIXA DIÁRIO"] = v

            v = self._read_value("CAIXA BANCO", self.V_CAIXA_BANCO)
            if v is not None:
                vals["CAIXA BANCO"] = v

            v = self._read_value("D. CRIANÇAS", self.V_CAIXA_CRIANCAS)
            if v is not None:
                vals["D. CRIANÇAS"] = v

            v = self._read_value("VERBO CAFE", self.V_CAIXA_CAFE)
            if v is not None:
                vals["VERBO CAFE"] = v

            v = self._read_value("VERBO SHOP", self.V_CAIXA_LIVRARIA)
            if v is not None:
                vals["VERBO SHOP"] = v

            if not vals:
                # nenhuma leitura funcionou => provavelmente não estamos no ecrã certo
                p = self._snap("caixas_no_values_found")
                log_kv(
                    log,
                    "Nenhum valor de Caixas/Bancos foi encontrado. Vou seguir sem atualizar a sheet.",
                    level=logging.WARNING,
                    url=getattr(self.a.driver, "current_url", None),
                    title=getattr(self.a.driver, "title", None),
                    screenshot=p,
                )
                return {}

            print("(6.2) Valores lidos do site:")
            for k, v in vals.items():
                print(f" - {k}: {v}")

            log_kv(log, "Valores Caixas/Bancos lidos.", level=logging.INFO, **vals)
            return vals
