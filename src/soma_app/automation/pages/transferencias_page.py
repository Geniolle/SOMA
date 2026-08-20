from __future__ import annotations

import logging
import time
import unicodedata
from typing import Any, List, Tuple

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from soma_app.automation.actions import Actions
from soma_app.config.locators import apply_locator_overrides
from soma_app.domain.models import ContaOrdemRow, format_amount_for_input
from soma_app.infra.trace import log_kv, step

log = logging.getLogger("soma_app.pages.transferencias")

Locator = Tuple[str, str]


class TransferenciasPage:
    """
    Fluxo de Transferência conforme SOMA.py:
      - Caixas/Bancos
      - Nova Transferência
      - Caixa Saída, Valor, Caixa Entrada, Data, Descrição
      - Salvar, OK, Voltar
    """

    MENU_CAIXAS_BANCOS_CANDIDATES: List[Locator] = []
    BTN_NOVA_TRANSFERENCIA_CANDIDATES: List[Locator] = []

    CAIXA_SAIDA = (By.XPATH, "")
    VALOR = (By.XPATH, "")
    CAIXA_ENTRADA = (By.XPATH, "")
    DATA = (By.XPATH, "")
    DESCRICAO = (By.XPATH, "")

    BTN_SALVAR = (By.XPATH, "")
    BTN_VOLTAR = (By.XPATH, "")

    OK_ALERT = (By.XPATH, "")
    SWAL_CONTAINER = (By.CLASS_NAME, "swal2-container")

    SELECT2_SEARCH = (By.XPATH, "")

    def __init__(self, actions: Actions, settings: Any):
        self.a = actions
        self.settings = settings
        self.home_url = (getattr(settings, "site_home_url", "") or "https://verbodavida.info/IVV/").strip()
        self.timeout = int(getattr(settings, "timeout_seconds", 20) or 20)
        apply_locator_overrides(self, "transferencias")

    def _dismiss_alerts(self) -> None:
        try:
            if self.a.exists(self.OK_ALERT, timeout_seconds=1):
                self.a.click_js(self.OK_ALERT)
        except Exception:
            pass
        try:
            if self.a.exists(self.SWAL_CONTAINER, timeout_seconds=1):
                self.a.wait_invisible(self.SWAL_CONTAINER, timeout_seconds=10)
        except Exception:
            pass

    def _click_any(self, candidates: List[Locator], timeout_seconds: int) -> Locator:
        loc = self.a.wait_any_present(candidates, timeout_seconds=timeout_seconds)
        self._dismiss_alerts()
        self.a.click_js(loc)
        self.a.wait_dom_ready(15)
        return loc

    def _goto_home(self) -> None:
        try:
            self.a.driver.get(self.home_url)
            self.a.wait_dom_ready(15)
            time.sleep(1)
        except Exception:
            pass

    @staticmethod
    def _norm_text(value: str) -> str:
        txt = unicodedata.normalize("NFKD", value or "")
        txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
        return " ".join(txt.lower().split())

    @staticmethod
    def _strip_accents(value: str) -> str:
        return "".join(ch for ch in unicodedata.normalize("NFKD", value or "") if not unicodedata.combining(ch))

    def _read_select2_selected_text(self, opener: Locator) -> str:
        try:
            host = self.a.driver.find_element(*opener)
        except Exception:
            return ""

        selectors = [
            "span.select2-selection__rendered",
            "span.select2-selection__choice",
            "span.select2-selection",
        ]

        for css in selectors:
            try:
                nodes = host.find_elements(By.CSS_SELECTOR, css)
            except Exception:
                nodes = []

            texts: List[str] = []
            for node in nodes:
                txt = (node.text or "").strip()
                if txt:
                    texts.append(txt)

            joined = " ".join(texts).strip()
            if joined:
                return joined

        try:
            return (host.get_attribute("textContent") or "").strip()
        except Exception:
            return ""

    def _collect_select2_options(self, value: str) -> List[str]:
        css_options = (By.CSS_SELECTOR, "span.select2-container--open li.select2-results__option")
        inp = self.a.wait_visible(self.SELECT2_SEARCH, timeout_seconds=30)

        queries: List[str] = []
        full = (value or "").strip()
        if full:
            queries.append(full)
            ascii_full = self._strip_accents(full).strip()
            if ascii_full and ascii_full not in queries:
                queries.append(ascii_full)
            first_token = full.split()[0].strip()
            if first_token and first_token not in queries:
                queries.append(first_token)

        last_sample: List[str] = []

        for query in queries:
            try:
                inp.clear()
            except Exception:
                pass
            inp.send_keys(query)
            time.sleep(0.8)

            try:
                self.a.driver.execute_script(
                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));",
                    inp,
                )
            except Exception:
                pass

            try:
                WebDriverWait(self.a.driver, 4).until(lambda d: len(d.find_elements(*css_options)) > 0)
            except Exception:
                continue

            options: List[str] = []
            for opt in self.a.driver.find_elements(*css_options):
                txt = (opt.text or "").strip()
                if not txt:
                    continue
                norm = self._norm_text(txt)
                if norm in {"no results found", "nenhum resultado encontrado", "sem resultados"}:
                    continue
                options.append(txt)

            last_sample = options
            if options:
                return options

        return last_sample

    def _select2_choose_verified(self, opener: Locator, value: str, *, row: ContaOrdemRow, field: str) -> None:
        v = (value or "").strip()
        if not v:
            raise ValueError(f"{field} vazio na sheet (linha {row.row_number}). Preenche a coluna correta.")

        self._dismiss_alerts()

        try:
            self.a.click(opener)
        except Exception:
            self.a.click_js(opener)

        inp = self.a.wait_visible(self.SELECT2_SEARCH, timeout_seconds=30)
        try:
            inp.clear()
        except Exception:
            pass
        inp.send_keys(v)

        WebDriverWait(self.a.driver, 5).until(
            EC.text_to_be_present_in_element_value(self.SELECT2_SEARCH, v)
        )

        options = self._collect_select2_options(v)
        if not options:
            raise RuntimeError(
                f"{field} não encontrou opção compatível com '{v}' (linha {row.row_number}). "
                "A lista retornou vazia."
            )

        want = self._norm_text(v)
        best_text = None
        for txt in options:
            norm = self._norm_text(txt)
            if norm == want or (want and want in norm) or (norm in want):
                best_text = txt
                break

        if best_text is None:
            best_text = options[0]

        css_options = (By.CSS_SELECTOR, "span.select2-container--open li.select2-results__option")
        best = None
        for opt in self.a.driver.find_elements(*css_options):
            if (opt.text or "").strip() == best_text:
                best = opt
                break

        if best is None:
            raise RuntimeError(
                f"{field} não encontrou opção compatível com '{v}' (linha {row.row_number}). "
                f"Amostra={options[:10]}"
            )

        try:
            self.a.driver.execute_script("arguments[0].click();", best)
        except Exception:
            best.click()

        time.sleep(0.8)

        try:
            txt = self._read_select2_selected_text(opener)
            if v.lower() not in txt.lower():
                time.sleep(1.2)
                txt = self._read_select2_selected_text(opener)
            if v.lower() not in txt.lower():
                raise RuntimeError(
                    f"{field} não foi selecionado (linha {row.row_number}). "
                    f"Esperado conter '{v}', mas ficou '{txt}'."
                )
        except Exception:
            p = self.a.screenshot(f"transfer_select2_fail_{field.lower().replace(' ', '_')}_row_{row.row_number}")
            log_kv(
                log,
                "Select2 não confirmou seleção.",
                level=logging.ERROR,
                field=field,
                row=row.row_number,
                value=v,
                url=self.a.driver.current_url,
                screenshot=p,
            )
            raise

    def open_new(self, row: ContaOrdemRow) -> None:
        print(f"\n[TRANSFERÊNCIA] Abrindo formulário | linha={row.row_number}")

        self._goto_home()

        with step(log, "transfer.open_menu", row=row.row_number):
            self._click_any(self.MENU_CAIXAS_BANCOS_CANDIDATES, timeout_seconds=max(60, self.timeout))
            time.sleep(5)

        with step(log, "transfer.open_new", row=row.row_number):
            self._click_any(self.BTN_NOVA_TRANSFERENCIA_CANDIDATES, timeout_seconds=max(60, self.timeout))
            time.sleep(2)
            self.a.wait_present(self.CAIXA_SAIDA, timeout_seconds=max(60, self.timeout))

    def fill_and_save(self, row: ContaOrdemRow) -> None:
        print(
            f"[TRANSFERÊNCIA] Preenchendo | linha={row.row_number} | "
            f"caixa_saida='{row.caixa_saida}' | caixa_entrada='{row.caixa}' | "
            f"valor='{row.importancia}' | data='{row.data_mov}'"
        )

        with step(log, "transfer.fill", row=row.row_number):
            self._select2_choose_verified(self.CAIXA_SAIDA, row.caixa_saida, row=row, field="CAIXA SAÍDA")
            self.a.type(self.VALOR, format_amount_for_input(row.importancia))
            time.sleep(0.5)
            self._select2_choose_verified(self.CAIXA_ENTRADA, row.caixa, row=row, field="CAIXA ENTRADA")
            self.a.type(self.DATA, row.data_mov)
            self.a.type(self.DESCRICAO, row.descricao_soma, clear=False)

        with step(log, "transfer.save", row=row.row_number):
            self._dismiss_alerts()
            self.a.click_js(self.BTN_SALVAR)
            time.sleep(1)
            self._dismiss_alerts()

        with step(log, "transfer.back", row=row.row_number):
            if self.a.exists(self.BTN_VOLTAR, timeout_seconds=15):
                self.a.click_js(self.BTN_VOLTAR)
                self.a.wait_dom_ready(15)
                time.sleep(1)

        print(f"[TRANSFERÊNCIA] Concluída | linha={row.row_number}")

    def run(self, row: ContaOrdemRow) -> str:
        with step(log, "transfer.run", row=row.row_number, tipo=row.tipo.value):
            self.open_new(row)
            self.fill_and_save(row)
            return "Transferido"
