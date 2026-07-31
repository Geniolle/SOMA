from __future__ import annotations

import logging
import time
import unicodedata
from typing import Any, List, Tuple

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from soma_app.automation.actions import Actions
from soma_app.config.fields import TransferenciaField, field_name
from soma_app.config.locators import apply_locator_overrides
from soma_app.config.urls import DEFAULT_SITE_HOME_URL
from soma_app.domain.models import ContaOrdemRow
from soma_app.infra.trace import log_kv, step

log = logging.getLogger("soma_app.pages.transferencias")

Locator = Tuple[str, str]


class TransferenciasPage:
    """
    Fluxo de Transferência conforme SOMA.py (fonte):
      - Caixas/Bancos
      - Nova Transferência
      - Caixa Saída, Valor, Caixa Entrada, Data, Descrição
      - Salvar, OK, Voltar
    """

    # Menu Caixas/Bancos (do SOMA.py)
    MENU_CAIXAS_BANCOS_CANDIDATES: List[Locator] = []

    # Botão Nova Transferência (do SOMA.py)
    BTN_NOVA_TRANSFERENCIA_CANDIDATES: List[Locator] = []

    # Campos Transferência (do SOMA.py)
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
        self.home_url = (getattr(settings, "site_home_url", "") or DEFAULT_SITE_HOME_URL).strip()
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
        # ajuda a evitar ficar em tela errada do SOMA
        try:
            self.a.driver.get(self.home_url)
            self.a.wait_dom_ready(15)
            time.sleep(0.2)
        except Exception:
            pass

    @staticmethod
    def _field(field: TransferenciaField | str) -> str:
        return field_name(field)

    @staticmethod
    def _select2_variants(value: str) -> List[str]:
        raw = " ".join((value or "").split())
        normalized = unicodedata.normalize("NFKD", raw)
        normalized = "".join(ch for ch in normalized if not unicodedata.combining(ch))
        normalized = " ".join(normalized.split())
        variants: List[str] = []
        for candidate in (raw, normalized, normalized.upper(), raw.upper()):
            candidate = " ".join(candidate.split())
            if candidate and candidate not in variants:
                variants.append(candidate)
        return variants

    def _select2_open_count(self) -> int:
        try:
            return len(self.a.driver.find_elements(By.CSS_SELECTOR, "span.select2-container--open"))
        except Exception:
            return 0

    def _select2_selected_text(self, opener: Locator) -> str:
        try:
            el = self.a.driver.find_element(*opener)
            return " ".join((getattr(el, "text", "") or "").split())
        except Exception:
            return ""

    @staticmethod
    def _select2_match_index(options: List[str], value: str) -> int | None:
        want = TransferenciasPage._select2_variants(value)
        want_norm = []
        for item in want:
            norm = unicodedata.normalize("NFKD", item).lower()
            norm = "".join(ch for ch in norm if not unicodedata.combining(ch))
            norm = " ".join(norm.split())
            if norm and norm not in want_norm:
                want_norm.append(norm)

        for idx, txt in enumerate(options):
            norm_txt = unicodedata.normalize("NFKD", txt or "").lower()
            norm_txt = "".join(ch for ch in norm_txt if not unicodedata.combining(ch))
            norm_txt = " ".join(norm_txt.split())
            if not norm_txt:
                continue
            for norm_want in want_norm:
                if norm_txt == norm_want or norm_want in norm_txt or norm_txt in norm_want:
                    return idx
        return None

    def _select2_direct_choose(self, opener: Locator, value: str, *, select_name: str | None = None) -> str:
        container = self.a.wait_present(opener, timeout_seconds=15)
        selectors = []
        if select_name:
            selectors.extend(
                [
                    f"select[name='{select_name}']",
                    f"select[name='{select_name}'].select2-hidden-accessible",
                ]
            )
        selectors.extend(["select.select2-hidden-accessible", "select"])

        select_el = None
        for css in selectors:
            try:
                if select_name and css.startswith("select[name="):
                    select_el = self.a.driver.find_element(By.CSS_SELECTOR, css)
                else:
                    select_el = container.find_element(By.CSS_SELECTOR, css)
                if select_el is not None:
                    break
            except Exception:
                continue
        if select_el is None:
            return ""

        options: List[str] = []
        try:
            options = [(" ".join((opt.text or "").split())) for opt in Select(select_el).options]
        except Exception:
            options = []

        idx = self._select2_match_index(options, value)
        if idx is None:
            return ""

        try:
            selected_text = self.a.driver.execute_script(
                """
                const select = arguments[0];
                const index = arguments[1];
                if (select && select.options && select.options[index]) {
                    const option = select.options[index];
                    if (window.jQuery) {
                        window.jQuery(select).val(option.value).trigger('change');
                        window.jQuery(select).trigger('change.select2');
                    } else {
                        select.selectedIndex = index;
                        select.dispatchEvent(new Event('input', { bubbles: true }));
                        select.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    select.dispatchEvent(new Event('input', { bubbles: true }));
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    return (select.options[select.selectedIndex] && select.options[select.selectedIndex].textContent) || option.textContent || '';
                }
                return '';
                """,
                select_el,
                idx,
            )
        except Exception:
            return ""

        return " ".join((selected_text or "").split())

    def _dump_select2_diagnostics(
        self,
        opener: Locator,
        *,
        suffix: str,
        row: ContaOrdemRow,
        value: str,
        select_name: str | None = None,
    ) -> None:
        name = f"transfer_select2_{suffix}_row_{row.row_number}"
        self.a.screenshot(name)
        self.a.dump_page_source(name)
        self.a.dump_locator_probe(name, [opener, self.SELECT2_SEARCH])
        try:
            container = self.a.wait_present(opener, timeout_seconds=5)
            selectors = []
            if select_name:
                selectors.extend(
                    [
                        (By.CSS_SELECTOR, f"select[name='{select_name}']"),
                        (By.CSS_SELECTOR, f"select[name='{select_name}'].select2-hidden-accessible"),
                    ]
                )
            selectors.extend([(By.CSS_SELECTOR, "select.select2-hidden-accessible"), (By.CSS_SELECTOR, "select")])
            select_el = None
            for loc in selectors:
                try:
                    if select_name and loc[1].startswith("select[name="):
                        select_el = self.a.driver.find_element(*loc)
                    else:
                        select_el = container.find_element(*loc)
                    break
                except Exception:
                    continue
            if select_el is None:
                return
            self.a.dump_locator_probe(name + "_select", [(By.CSS_SELECTOR, "select.select2-hidden-accessible")])
            try:
                opts = [(" ".join((opt.text or "").split())) for opt in Select(select_el).options]
            except Exception:
                opts = []
            log_kv(
                log,
                "Select2 diagnostics.",
                level=logging.WARNING,
                row=row.row_number,
                field="CAIXA_SAIDA",
                value=value,
                options=opts[:10],
                current=self._select2_selected_text(opener),
            )
        except Exception:
            pass

    def _select2_choose_verified(self, opener: Locator, value: str, *, row: ContaOrdemRow, field: TransferenciaField | str) -> None:
        v = (value or "").strip()
        field_key = self._field(field)
        if not v:
            raise ValueError(f"{field_key} vazio na sheet (linha {row.row_number}). Preenche a coluna correta (ex.: CAIXA SAIDA).")

        select_name = None
        if field_key == self._field(TransferenciaField.CAIXA_SAIDA):
            select_name = "id_caixa_origem"
        elif field_key == self._field(TransferenciaField.CAIXA_ENTRADA):
            select_name = "id_caixa_destino"

        self._dismiss_alerts()

        # primeiro tenta selecionar o <select> real por baixo do select2.
        chosen_direct = self._select2_direct_choose(opener, v, select_name=select_name)
        if chosen_direct:
            want_norm = unicodedata.normalize("NFKD", v).lower()
            want_norm = "".join(ch for ch in want_norm if not unicodedata.combining(ch))
            chosen_norm = unicodedata.normalize("NFKD", chosen_direct).lower()
            chosen_norm = "".join(ch for ch in chosen_norm if not unicodedata.combining(ch))
            if want_norm and (want_norm in chosen_norm or chosen_norm in want_norm):
                return

        # abre o select2
        try:
            self.a.click(opener)
        except Exception:
            self.a.click_js(opener)

        last_error: Exception | None = None

        for attempt_value in self._select2_variants(v):
            inp = self.a.wait_visible(self.SELECT2_SEARCH, timeout_seconds=30)
            inp.clear()
            inp.send_keys(attempt_value)

            try:
                WebDriverWait(self.a.driver, 5).until(
                    EC.text_to_be_present_in_element_value(self.SELECT2_SEARCH, attempt_value)
                )
            except Exception:
                pass

            try:
                options = self.a.driver.find_elements(By.CSS_SELECTOR, "li.select2-results__option")
            except Exception:
                options = []

            want = unicodedata.normalize("NFKD", attempt_value).lower()
            want = "".join(ch for ch in want if not unicodedata.combining(ch))
            chosen = False

            for opt in options:
                txt = " ".join((opt.text or "").split())
                if not txt or "no results found" in txt.lower():
                    continue
                norm_txt = unicodedata.normalize("NFKD", txt).lower()
                norm_txt = "".join(ch for ch in norm_txt if not unicodedata.combining(ch))
                if norm_txt == want or want in norm_txt:
                    try:
                        self.a.driver.execute_script("arguments[0].click();", opt)
                    except Exception:
                        opt.click()
                    chosen = True
                    break

            if not chosen:
                inp.send_keys(Keys.ENTER)

            time.sleep(0.3)

            try:
                WebDriverWait(self.a.driver, 5).until(lambda _driver: self._select2_open_count() == 0)
            except Exception as exc:
                last_error = exc
                continue

            txt = self._select2_selected_text(opener)
            norm_txt = unicodedata.normalize("NFKD", txt).lower()
            norm_txt = "".join(ch for ch in norm_txt if not unicodedata.combining(ch))
            if want and (want in norm_txt or norm_txt in want):
                return

        p = self.a.screenshot(f"transfer_select2_fail_{field_key.lower()}_row_{row.row_number}")
        try:
            self._dump_select2_diagnostics(opener, suffix=f"fail_{field_key.lower()}", row=row, value=v, select_name=select_name)
        except Exception:
            pass
        log_kv(
            log,
            "Select2 não confirmou seleção.",
            level=logging.ERROR,
            field=field_key,
            row=row.row_number,
            value=v,
            url=self.a.driver.current_url,
            screenshot=p,
            err=type(last_error).__name__ if last_error else "SelectionMismatch",
        )
        raise RuntimeError(f"{field_key} não foi selecionado (linha {row.row_number}). Esperado conter '{v}'.")

    def open_new(self, row: ContaOrdemRow) -> None:
        # console (para aparecer sempre)
        print(f"\n[TRANSFERÊNCIA] Abrindo formulário | linha={row.row_number}")

        self._goto_home()

        with step(log, "transfer.open_menu", row=row.row_number):
            self._click_any(self.MENU_CAIXAS_BANCOS_CANDIDATES, timeout_seconds=max(60, self.timeout))

        with step(log, "transfer.open_new", row=row.row_number):
            # tenta abrir “Nova Transferência”
            self._click_any(self.BTN_NOVA_TRANSFERENCIA_CANDIDATES, timeout_seconds=max(60, self.timeout))

            # espera aparecer o primeiro campo do form
            self.a.wait_present(self.CAIXA_SAIDA, timeout_seconds=max(60, self.timeout))

    def fill_and_save(self, row: ContaOrdemRow) -> None:
        # console (para aparecer sempre)
        print(
            f"[TRANSFERÊNCIA] Preenchendo | linha={row.row_number} | "
            f"caixa_saida='{row.caixa_saida}' | caixa_entrada='{row.caixa}' | valor='{row.importancia}' | data='{row.data_mov}'"
        )

        with step(log, "transfer.fill", row=row.row_number):
            self._select2_choose_verified(self.CAIXA_SAIDA, row.caixa_saida, row=row, field=self._field(TransferenciaField.CAIXA_SAIDA))
            self.a.type(self.VALOR, str(row.importancia))
            self._select2_choose_verified(self.CAIXA_ENTRADA, row.caixa, row=row, field=self._field(TransferenciaField.CAIXA_ENTRADA))
            self.a.type(self.DATA, row.data_mov)
            self.a.type(self.DESCRICAO, row.descricao_soma, clear=False)

        with step(log, "transfer.save", row=row.row_number):
            self._dismiss_alerts()
            self.a.click_js(self.BTN_SALVAR)
            self._dismiss_alerts()

        with step(log, "transfer.back", row=row.row_number):
            if self.a.exists(self.BTN_VOLTAR, timeout_seconds=15):
                self.a.click_js(self.BTN_VOLTAR)
                self.a.wait_dom_ready(15)

        print(f"[TRANSFERÊNCIA] Concluída | linha={row.row_number}")

    def run(self, row: ContaOrdemRow) -> str:
        with step(log, "transfer.run", row=row.row_number, tipo=row.tipo.value):
            self.open_new(row)
            self.fill_and_save(row)
            return "Transferido"
