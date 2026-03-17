from __future__ import annotations

import logging
import time
from typing import Any, List, Tuple

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from soma_app.automation.actions import Actions
from soma_app.domain.models import ContaOrdemRow
from soma_app.infra.trace import step, log_kv

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
    MENU_CAIXAS_BANCOS_CANDIDATES: List[Locator] = [
        (By.XPATH, "/html/body/div[2]/div/div[1]/div[2]/div/div/div/div[4]/div/div/span"),
        (By.XPATH, "//*[self::span or self::a or self::button][contains(.,'Caixas') and contains(.,'Bancos')]"),
        (By.XPATH, "//*[self::span or self::a or self::button][contains(.,'Caixas') and contains(.,'bancos')]"),
    ]

    # Botão Nova Transferência (do SOMA.py)
    BTN_NOVA_TRANSFERENCIA_CANDIDATES: List[Locator] = [
        (By.XPATH, "/html/body/div[2]/div/div[2]/a"),
        (By.XPATH, "//*[self::a or self::button][contains(.,'Nova') and (contains(.,'Transfer') or contains(.,'transfer'))]"),
    ]

    # Campos Transferência (do SOMA.py)
    CAIXA_SAIDA = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[2]/div")
    VALOR = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[4]/div/input")
    CAIXA_ENTRADA = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[5]/div")
    DATA = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[8]/div/div/input")
    DESCRICAO = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[9]/div/textarea")

    BTN_SALVAR = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[10]/div/button[1]")
    BTN_VOLTAR = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[10]/div/button[2]")

    OK_ALERT = (By.XPATH, "/html/body/div[5]/div/button[1]")
    SWAL_CONTAINER = (By.CLASS_NAME, "swal2-container")

    SELECT2_SEARCH = (By.XPATH, "/html/body/span/span/span[1]/input")

    def __init__(self, actions: Actions, settings: Any):
        self.a = actions
        self.settings = settings
        self.home_url = (getattr(settings, "site_home_url", "") or "https://verbodavida.info/IVV/").strip()
        self.timeout = int(getattr(settings, "timeout_seconds", 20) or 20)

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
            time.sleep(1)
        except Exception:
            pass

    def _select2_choose_verified(self, opener: Locator, value: str, *, row: ContaOrdemRow, field: str) -> None:
        v = (value or "").strip()
        if not v:
            raise ValueError(f"{field} vazio na sheet (linha {row.row_number}). Preenche a coluna correta (ex.: CAIXA SAIDA).")

        self._dismiss_alerts()

        # abre o select2
        try:
            self.a.click(opener)
        except Exception:
            self.a.click_js(opener)

        # espera o input do select2
        inp = self.a.wait_visible(self.SELECT2_SEARCH, timeout_seconds=30)
        inp.clear()
        inp.send_keys(v)

        # igual ao SOMA.py: espera o valor aparecer no input antes do ENTER
        WebDriverWait(self.a.driver, 5).until(
            EC.text_to_be_present_in_element_value(self.SELECT2_SEARCH, v)
        )
        inp.send_keys(Keys.ENTER)
        time.sleep(0.8)

        # verificação “best effort”: o texto do opener deve conter o valor selecionado
        try:
            el = self.a.driver.find_element(*opener)
            txt = (el.text or "").strip()
            if v.lower() not in txt.lower():
                # às vezes demora renderizar, tenta mais um pouco
                time.sleep(1.2)
                el = self.a.driver.find_element(*opener)
                txt = (el.text or "").strip()
            if v.lower() not in txt.lower():
                raise RuntimeError(f"{field} não foi selecionado (linha {row.row_number}). Esperado conter '{v}', mas ficou '{txt}'.")
        except Exception as e:
            p = self.a.screenshot(f"transfer_select2_fail_{field.lower().replace(' ','_')}_row_{row.row_number}")
            log_kv(log, logging.ERROR, "Select2 não confirmou seleção.", field=field, row=row.row_number, value=v, url=self.a.driver.current_url, screenshot=p)
            raise

    def open_new(self, row: ContaOrdemRow) -> None:
        # console (para aparecer sempre)
        print(f"\n[TRANSFERÊNCIA] Abrindo formulário | linha={row.row_number}")

        self._goto_home()

        with step(log, "transfer.open_menu", row=row.row_number):
            self._click_any(self.MENU_CAIXAS_BANCOS_CANDIDATES, timeout_seconds=max(60, self.timeout))
            time.sleep(5)  # igual ao SOMA.py

        with step(log, "transfer.open_new", row=row.row_number):
            # tenta abrir “Nova Transferência”
            self._click_any(self.BTN_NOVA_TRANSFERENCIA_CANDIDATES, timeout_seconds=max(60, self.timeout))
            time.sleep(2)

            # espera aparecer o primeiro campo do form
            self.a.wait_present(self.CAIXA_SAIDA, timeout_seconds=max(60, self.timeout))

    def fill_and_save(self, row: ContaOrdemRow) -> None:
        # console (para aparecer sempre)
        print(
            f"[TRANSFERÊNCIA] Preenchendo | linha={row.row_number} | "
            f"caixa_saida='{row.caixa_saida}' | caixa_entrada='{row.caixa}' | valor='{row.importancia}' | data='{row.data_mov}'"
        )

        with step(log, "transfer.fill", row=row.row_number):
            self._select2_choose_verified(self.CAIXA_SAIDA, row.caixa_saida, row=row, field="CAIXA SAÍDA")
            self.a.type(self.VALOR, str(row.importancia))
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