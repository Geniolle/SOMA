from __future__ import annotations

import logging
import time
import unicodedata
from typing import Any, Callable, Iterable, List, Tuple

from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from soma_app.automation.actions import Actions
from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra.trace import step, log_kv

log = logging.getLogger("soma_app.pages.entradas_saidas")


class EntradasSaidasPage:
    # =========
    # MENU / NOVO (robusto por candidatos)
    # =========
    MENU_ENTRADAS_SAIDAS_CANDIDATES = [
        (
            By.XPATH,
            "//*[self::a or self::button or self::span]"
            "[contains(.,'Entradas') and "
            "(contains(.,'Saídas') or contains(.,'Saidas') or contains(.,'SaÃdas') or contains(.,'saÃdas') or contains(.,'SaÃdas'))]",
        ),
        (
            By.XPATH,
            "//*[contains(.,'Entradas') and "
            "(contains(.,'Saídas') or contains(.,'Saidas') or contains(.,'SaÃdas') or contains(.,'saÃdas'))]",
        ),
    ]

    BTN_NOVA_CANDIDATES = [
        (By.XPATH, "//a[contains(@class,'btn') and contains(.,'Nova') and contains(.,'Entrada')]"),
        (By.XPATH, "//button[contains(@class,'btn') and contains(.,'Nova') and contains(.,'Entrada')]"),
        (By.XPATH, "//*[self::a or self::button][contains(.,'Nova') and contains(.,'Entrada')]"),
        (By.XPATH, '//a[contains(@class, "btn btn-primary") and contains(., "Nova Entrada/Saída")]'),
    ]

    # =========
    # RADIO TIPO
    # =========
    RADIO_SAIDA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[1]/input")
    RADIO_ENTRADA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/div[2]/input")

    FORM_CONTAINER = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form")

    RADIO_ENTRADA_CANDIDATES = [
        RADIO_ENTRADA,
        (By.XPATH, "//input[@type='radio' and contains(@value,'Entrada')]"),
        (By.XPATH, "//input[@type='radio' and contains(@value,'entrada')]"),
        (By.XPATH, "//input[@type='radio' and (contains(@id,'Entrada') or contains(@name,'Entrada'))]"),
        (By.XPATH, "//input[@type='radio' and (contains(@id,'entrada') or contains(@name,'entrada'))]"),
    ]

    RADIO_SAIDA_CANDIDATES = [
        RADIO_SAIDA,
        (By.XPATH, "//input[@type='radio' and (contains(@value,'Saída') or contains(@value,'Saida'))]"),
        (By.XPATH, "//input[@type='radio' and (contains(@value,'saída') or contains(@value,'saida'))]"),
        (
            By.XPATH,
            "//input[@type='radio' and (contains(@id,'Saída') or contains(@id,'Saida') or contains(@name,'Saída') or contains(@name,'Saida'))]",
        ),
        (
            By.XPATH,
            "//input[@type='radio' and (contains(@id,'saída') or contains(@id,'saida') or contains(@name,'saída') or contains(@name,'saida'))]",
        ),
    ]

    RADIO_ANY_CANDIDATES = RADIO_SAIDA_CANDIDATES + RADIO_ENTRADA_CANDIDATES

    # =========
    # CAMPOS COMUNS
    # =========
    PLANO_CONTA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[6]/div/span/span[1]/span/span[1]")
    CENTRO_CUSTO = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span")
    DESCRICAO = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/input")
    VALOR = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[12]/div/input")
    OBS = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[21]/div/textarea")

    # =========
    # CAMPOS ENTRADA
    # =========
    DATA_ENTRADA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[11]/div/div/input")
    FORMA_PAGAMENTO_ENTRADA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[16]/div/span/span[1]/span")
    CAIXA_ENTRADA_CONTAINER = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[20]/div")

    # =========
    # CAMPOS SAÍDA
    # =========
    DATA_VENCIMENTO_SAIDA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/div/input")

    # =========
    # BOTÕES (form)
    # =========
    BTN_SALVAR_FORM = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[30]/div/button[1]")
    BTN_VOLTAR = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[30]/div/button[2]")

    FORM_READY_CANDIDATES = [
        FORM_CONTAINER,
        PLANO_CONTA,
        DESCRICAO,
        BTN_SALVAR_FORM,
    ]

    # =========
    # ALERTAS / OVERLAY
    # =========
    OK_ALERT = (By.XPATH, "/html/body/div[5]/div/button[1]")
    SWAL_CONTAINER = (By.CLASS_NAME, "swal2-container")

    # =========
    # PAGAMENTO / BAIXA
    # =========
    BTN_REALIZAR_PAGAMENTO = (By.XPATH, "/html/body/ul/li/div/div[2]/button[1]")

    BTN_INSERIR_PAGAMENTO_SAIDA = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[29]/div/div[2]/a")
    DATA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input")
    FORMA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select")
    CAIXA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[4]/div[2]/div/select")
    NUM_DOCUMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[3]/div/input")
    BTN_SALVAR_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[3]/button")

    BTN_INSERIR_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button")
    DATA_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input")
    POPUP_CLICK_CANDIDATES = [
        (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[2]"),
        (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/form/div[3]"),
    ]
    BTN_SALVAR_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/form/div[3]/button")

    # =========
    # PESQUISA DOC SOMA
    # =========
    PESQ_DESCRICAO = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[3]/div/input")
    RADIO_PERIODO = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/div/input")
    RADIO_DATA_PAGAMENTO = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[8]/div/div[2]/input")
    DATA_INI = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[1]/div/input")
    DATA_FIM = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[9]/div/div[3]/div/input")
    BTN_PESQUISAR = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[10]/div/button")
    RESULT_DOC = (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/div/div/table/tbody/tr/td[3]")

    NO_RESULTS_CANDIDATES = [
        (By.XPATH, "//*[@class='dataTables_empty']"),
        (By.XPATH, "//*[contains(.,'Nenhum registo') or contains(.,'Nenhum registro') or contains(.,'Sem resultados')]"),
        (By.XPATH, "//*[contains(.,'No matching records') or contains(.,'No data available') or contains(.,'Nothing found')]"),
    ]

    # =========
    # DADOS DOC
    # =========
    DADOS_DOC_CELL = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[4]")

    def __init__(self, actions: Actions, settings: Any):
        self.a = actions
        self.settings = settings
        self.base_ivv = (getattr(settings, "site_home_url", "") or "https://verbodavida.info/IVV/").rstrip("/") + "/"

        self.strict_caixa = bool(getattr(settings, "STRICT_CAIXA_MATCH", True))

        self.caixa_validate_sleep = float(getattr(settings, "CAIXA_VALIDATE_SLEEP", 1.2))
        self.caixa_validate_retries = int(getattr(settings, "CAIXA_VALIDATE_RETRIES", 3))
        self.caixa_stable_checks = int(getattr(settings, "CAIXA_STABLE_CHECKS", 2))
        self.caixa_stable_interval = float(getattr(settings, "CAIXA_STABLE_INTERVAL", 0.4))

    # -----------------------
    # helpers de log/console
    # -----------------------
    @staticmethod
    def _safe(v: Any, max_len: int = 160) -> str:
        s = "" if v is None else str(v)
        s = " ".join(s.split())
        return s[:max_len] + ("…" if len(s) > max_len else "")

    @staticmethod
    def _norm(s: str) -> str:
        s2 = unicodedata.normalize("NFKD", (s or ""))
        s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
        return " ".join(s2.strip().lower().split())

    def _emit(self, msg: str, level: int = logging.INFO, **kv: Any) -> None:
        print(msg)
        try:
            log_kv(log, level, msg, **kv)
        except Exception:
            log.log(level, msg)

    def _exists_any(self, locators: Iterable[Tuple[str, str]], timeout_seconds: int = 1) -> bool:
        for loc in locators:
            try:
                if self.a.exists(loc, timeout_seconds=timeout_seconds):
                    return True
            except Exception:
                pass
        return False

    def _input_value(self, locator: Tuple[str, str]) -> str:
        try:
            el = self.a.driver.find_element(*locator)
            return (el.get_attribute("value") or "").strip()
        except Exception:
            return ""

    def _select_first_text(self, select_el) -> str:
        try:
            return (Select(select_el).first_selected_option.text or "").strip()
        except Exception:
            return ""

    def _match_ok(self, desired: str, actual: str) -> bool:
        d = self._norm(desired)
        a = self._norm(actual)
        if not d or not a:
            return False
        return d == a or (d in a) or (a in d)

    def _wait_select_ready(self, resolve_select: Callable[[], Any], timeout_seconds: int = 10, min_options: int = 2) -> None:
        t0 = time.time()
        while time.time() - t0 <= timeout_seconds:
            try:
                sel_el = resolve_select()
                opts = Select(sel_el).options
                if len(opts) >= min_options:
                    return
            except StaleElementReferenceException:
                pass
            except Exception:
                pass
            time.sleep(0.25)

    def _select_best_effort(
        self,
        select_el,
        desired_text: str,
        *,
        row: ContaOrdemRow,
        field: str,
        strict: bool,
    ) -> str:
        desired = (desired_text or "").strip()
        if not desired:
            raise RuntimeError(f"{field}: valor desejado vazio (row={row.row_number}).")

        sel = Select(select_el)
        options = [(o.text or "").strip() for o in sel.options]
        options_norm = [self._norm(t) for t in options]
        want_norm = self._norm(desired)

        try:
            sel.select_by_visible_text(desired)
            return self._select_first_text(select_el)
        except Exception:
            pass

        for idx, on in enumerate(options_norm):
            if on and on == want_norm:
                sel.select_by_index(idx)
                return self._select_first_text(select_el)

        candidates: List[int] = []
        for idx, on in enumerate(options_norm):
            if on and want_norm and (want_norm in on or on in want_norm):
                candidates.append(idx)

        if candidates:
            candidates.sort(key=lambda i: len(options_norm[i]), reverse=True)
            sel.select_by_index(candidates[0])
            chosen = self._select_first_text(select_el)
            if strict and not self._match_ok(desired, chosen):
                raise RuntimeError(
                    f"{field}: seleção ambígua/incorreta. desejado='{desired}' | selecionado='{chosen}' | opções={options}"
                )
            return chosen

        chosen = self._select_first_text(select_el)
        if strict:
            raise RuntimeError(
                f"{field}: não encontrei opção para '{desired}'. selecionado_atual='{chosen}' | opções={options}"
            )
        return chosen

    def _select_with_sleep_validation(
        self,
        resolve_select: Callable[[], Any],
        desired_text: str,
        *,
        row: ContaOrdemRow,
        field: str,
        strict: bool,
    ) -> str:
        self._wait_select_ready(resolve_select, timeout_seconds=10, min_options=2)

        desired = (desired_text or "").strip()
        last_seen = ""
        for attempt in range(1, self.caixa_validate_retries + 1):
            try:
                sel_el = resolve_select()

                chosen_now = self._select_best_effort(
                    sel_el,
                    desired,
                    row=row,
                    field=field,
                    strict=False,
                )

                time.sleep(self.caixa_validate_sleep)

                stable = True
                prev = None
                for _ in range(max(1, self.caixa_stable_checks)):
                    try:
                        cur = self._select_first_text(resolve_select())
                    except StaleElementReferenceException:
                        stable = False
                        break
                    if prev is None:
                        prev = cur
                    elif cur != prev:
                        stable = False
                        break
                    time.sleep(self.caixa_stable_interval)

                try:
                    current = self._select_first_text(resolve_select())
                except StaleElementReferenceException:
                    current = ""

                last_seen = current or chosen_now

                if stable and self._match_ok(desired, last_seen):
                    return last_seen

                self._emit(
                    f"Aviso: {field} alterou após seleção. tentativa {attempt}/{self.caixa_validate_retries} | "
                    f"esperado='{desired}' | agora='{last_seen}'. Reaplicando...",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )

            except StaleElementReferenceException:
                self._emit(
                    f"Aviso: {field} ficou stale durante a seleção. tentativa {attempt}/{self.caixa_validate_retries}. Reaplicando...",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                time.sleep(0.3)
                continue

        if strict:
            raise RuntimeError(f"{field}: não manteve a seleção. esperado='{desired}' | final='{last_seen}'")
        return last_seen

    # -----------------------
    # helpers UI
    # -----------------------
    def _close_datepicker(self) -> None:
        try:
            self.a.driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass

    def _dismiss_overlays(self) -> None:
        try:
            if self.a.exists(self.OK_ALERT, timeout_seconds=1):
                self.a.click_js(self.OK_ALERT)
                self._emit("Botão 'OK' clicado com sucesso!")
        except Exception:
            pass

        try:
            if self.a.exists(self.SWAL_CONTAINER, timeout_seconds=1):
                try:
                    self.a.wait_invisible(self.SWAL_CONTAINER, timeout_seconds=10)
                except Exception:
                    self._emit("Aviso: O modal SweetAlert ainda está visível após o tempo de espera.", level=logging.WARNING)
        except Exception:
            pass

    def _click_menu_entradas_saidas(self, timeout_seconds: int = 60) -> None:
        loc = self.a.wait_any_present(self.MENU_ENTRADAS_SAIDAS_CANDIDATES, timeout_seconds=timeout_seconds)
        self.a.click_js(loc)
        self.a.wait_dom_ready(15)
        self._emit("Clicou na opção 'Entradas/Saídas' com sucesso!")

    def _click_nova(self, timeout_seconds: int = 30) -> None:
        loc = self.a.wait_any_present(self.BTN_NOVA_CANDIDATES, timeout_seconds=timeout_seconds)
        self.a.click_js(loc)
        self.a.wait_dom_ready(15)
        self._emit("Clicou no botão 'Nova Entrada/Saída' com sucesso!")

    def _ensure_pesquisa_visivel(self, row: ContaOrdemRow) -> None:
        if self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=2):
            return

        self._dismiss_overlays()
        self._close_datepicker()

        with step(log, "entradas_saidas.ensure_list", row=row.row_number, tipo=row.tipo.value):
            self._click_menu_entradas_saidas(timeout_seconds=60)
            self.a.wait_present(self.PESQ_DESCRICAO, timeout_seconds=30)

    # -----------------------
    # navegação principal
    # -----------------------
    def _wait_new_form_ready(self, row: ContaOrdemRow, timeout_seconds: int = 60) -> None:
        last_err: Exception | None = None

        for attempt in (1, 2):
            try:
                self.a.wait_any_present(self.FORM_READY_CANDIDATES, timeout_seconds=timeout_seconds)
                time.sleep(0.8)
                self.a.wait_any_present(self.RADIO_ANY_CANDIDATES, timeout_seconds=min(30, timeout_seconds))
                return
            except Exception as e:
                last_err = e
                try:
                    p = self.a.screenshot(f"entradas_saidas_new_form_row_{row.row_number}_try_{attempt}")
                    log_kv(
                        log,
                        logging.ERROR,
                        "Form 'Nova Entrada/Saída' não ficou pronto. Vou tentar novamente.",
                        row=row.row_number,
                        tipo=row.tipo.value,
                        attempt=attempt,
                        url=getattr(self.a.driver, "current_url", ""),
                        title=getattr(self.a.driver, "title", ""),
                        screenshot=p,
                    )
                except Exception:
                    pass

                try:
                    self._dismiss_overlays()
                    self._close_datepicker()
                    self._click_menu_entradas_saidas(timeout_seconds=60)
                    time.sleep(0.8)
                    self._click_nova(timeout_seconds=30)
                    self.a.wait_dom_ready(15)
                except Exception:
                    pass

        raise TimeoutException(
            f"Timeout ao abrir formulário 'Nova Entrada/Saída' (linha {row.row_number}). last_err={last_err}"
        )

    def _open_new(self, row: ContaOrdemRow) -> None:
        self._dismiss_overlays()
        self._close_datepicker()

        with step(log, "entradas_saidas.open_menu", row=row.row_number, tipo=row.tipo.value):
            self._click_menu_entradas_saidas(timeout_seconds=60)

        with step(log, "entradas_saidas.open_new_form", row=row.row_number, tipo=row.tipo.value):
            self._click_nova(timeout_seconds=30)
            self.a.wait_dom_ready(15)
            self._wait_new_form_ready(row, timeout_seconds=60)
            time.sleep(1)

    def _choose_tipo(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.choose_tipo", row=row.row_number, tipo=row.tipo.value):
            if row.tipo == TipoMovimento.SAIDA:
                loc = self.a.wait_any_present(self.RADIO_SAIDA_CANDIDATES, timeout_seconds=30)
                self.a.click_js(loc)
                self._emit("Selecionado o botão para o processo de 'Saída'")
            else:
                loc = self.a.wait_any_present(self.RADIO_ENTRADA_CANDIDATES, timeout_seconds=30)
                self.a.click_js(loc)
                self._emit("Selecionado o botão para o processo de 'Entrada'")
            time.sleep(1)

    # -----------------------
    # preenchimento
    # -----------------------
    def _fill_common(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.plano_conta", row=row.row_number, tipo=row.tipo.value, field="PLANO_CONTA"):
            self.a.select2_choose(self.PLANO_CONTA, row.plano_conta)
            self._emit(f"Plano de conta preenchido com sucesso: {row.plano_conta}", row=row.row_number, tipo=row.tipo.value)

        with step(log, "entradas_saidas.fill.centro_custo", row=row.row_number, tipo=row.tipo.value, field="CENTRO_CUSTO"):
            self.a.select2_choose(self.CENTRO_CUSTO, row.centro_custo)
            self._emit(f"Centro de custo preenchido com sucesso: {row.centro_custo}", row=row.row_number, tipo=row.tipo.value)

        with step(log, "entradas_saidas.fill.descricao", row=row.row_number, tipo=row.tipo.value, field="DESCRICAO"):
            self.a.type(self.DESCRICAO, row.descricao_soma, clear=False)
            v = self._input_value(self.DESCRICAO) or row.descricao_soma
            self._emit(f"Descrição preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

        with step(log, "entradas_saidas.fill.valor", row=row.row_number, tipo=row.tipo.value, field="VALOR"):
            self.a.type(self.VALOR, str(row.importancia))
            try:
                self.a.driver.find_element(*self.VALOR).send_keys(Keys.ENTER)
            except Exception:
                pass
            v = self._input_value(self.VALOR) or str(row.importancia)
            self._emit(f"Valor preenchido com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

        with step(log, "entradas_saidas.fill.obs", row=row.row_number, tipo=row.tipo.value, field="OBS"):
            self.a.click_js(self.OBS)
            self.a.type(self.OBS, row.descricao_soma, clear=False)
            v = self._input_value(self.OBS) or row.descricao_soma
            self._emit(f"Descrição preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

    def _fill_entrada_sem_caixa(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.data_entrada", row=row.row_number, tipo=row.tipo.value, field="DATA_ENTRADA"):
            self.a.type(self.DATA_ENTRADA, row.data_mov)
            self._close_datepicker()
            v = self._input_value(self.DATA_ENTRADA) or row.data_mov
            self._emit(f"Data preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

        with step(
            log,
            "entradas_saidas.fill.forma_pagamento_entrada",
            row=row.row_number,
            tipo=row.tipo.value,
            field="FORMA_PAGAMENTO",
        ):
            self.a.select2_choose(self.FORMA_PAGAMENTO_ENTRADA, row.forma_pagamento)
            self._emit(f"Forma de pagamento selecionada com sucesso: {row.forma_pagamento}", row=row.row_number, tipo=row.tipo.value)

    def _resolve_caixa_entrada_select(self) -> Any:
        # re-localiza sempre + clica no container (como no SOMA.py) para forçar carregamento
        self.a.wait_present(self.CAIXA_ENTRADA_CONTAINER, timeout_seconds=10)
        container = self.a.driver.find_element(*self.CAIXA_ENTRADA_CONTAINER)
        try:
            container.click()
        except Exception:
            pass
        # pequeno respiro para o DOM refletir a abertura/refresh
        time.sleep(0.2)
        return container.find_element(By.TAG_NAME, "select")

    def _fill_caixa_entrada_ultima(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.caixa_entrada_last", row=row.row_number, tipo=row.tipo.value, field="CAIXA"):
            if self.a.exists(self.CAIXA_ENTRADA_CONTAINER, timeout_seconds=2):
                # respiro extra pós forma_pagamento (evita stale)
                time.sleep(0.5)

                chosen = self._select_with_sleep_validation(
                    self._resolve_caixa_entrada_select,
                    row.caixa,
                    row=row,
                    field="CAIXA_ENTRADA",
                    strict=self.strict_caixa,
                )
                self._emit(f"Caixa selecionada com sucesso: {chosen}", row=row.row_number, tipo=row.tipo.value)

                if self.strict_caixa and not self._match_ok(row.caixa, chosen):
                    raise RuntimeError(f"CAIXA_ENTRADA incorreta. esperado='{row.caixa}' | selecionado='{chosen}'")
            else:
                self._emit("Aviso: Campo de Caixa (Entrada) não encontrado no formulário.", level=logging.WARNING)

    def _fill_saida(self, row: ContaOrdemRow) -> None:
        with step(
            log,
            "entradas_saidas.fill.data_vencimento_saida",
            row=row.row_number,
            tipo=row.tipo.value,
            field="DATA_VENCIMENTO",
        ):
            self.a.type(self.DATA_VENCIMENTO_SAIDA, row.data_mov)
            self._close_datepicker()
            v = self._input_value(self.DATA_VENCIMENTO_SAIDA) or row.data_mov
            self._emit(f"Data vencimento preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

    def _save_form_if_present(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.save_form_best_effort", row=row.row_number, tipo=row.tipo.value):
            try:
                self._dismiss_overlays()
                self._close_datepicker()
                if self.a.exists(self.BTN_SALVAR_FORM, timeout_seconds=2):
                    self.a.click_js(self.BTN_SALVAR_FORM)
                    time.sleep(2)
                    self._dismiss_overlays()
                    self._emit("Campos principais preenchidos com sucesso")
            except Exception:
                pass

    # -----------------------
    # pagamento/baixa
    # -----------------------
    def _realizar_pagamento(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.realizar_pagamento", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays()
            self.a.click_js(self.BTN_REALIZAR_PAGAMENTO)
            time.sleep(2)
            self._emit("Realizar pagamento salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)

    def _resolve_caixa_pagamento_modal_select(self) -> Any:
        return self.a.driver.find_element(*self.CAIXA_PAGAMENTO_MODAL)

    def _pagamento_saida_modal(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.pagamento_saida_modal", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays()

            self.a.click_js(self.BTN_INSERIR_PAGAMENTO_SAIDA)
            time.sleep(1)
            self._emit("Inserir pagamento salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)

            self.a.type(self.DATA_PAGAMENTO_MODAL, row.data_mov)
            time.sleep(0.5)
            v = self._input_value(self.DATA_PAGAMENTO_MODAL) or row.data_mov
            self._emit(f"Data início preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

            fp_el = self.a.driver.find_element(*self.FORMA_PAGAMENTO_MODAL)
            chosen_fp = self._select_best_effort(
                fp_el,
                row.forma_pagamento,
                row=row,
                field="FORMA_PAGAMENTO_MODAL",
                strict=True,
            )
            self._emit(f"Forma de pagamento selecionada com sucesso: {chosen_fp}", row=row.row_number, tipo=row.tipo.value)
            time.sleep(0.5)

            if (row.forma_pagamento or "").strip().upper() == "TRANSFERÊNCIA BANCÁRIA":
                self._emit(
                    "A forma de pagamento é TRANSFERÊNCIA BANCÁRIA. atualizar campo Nº Documento...",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                try:
                    self.a.type(self.NUM_DOCUMENTO_MODAL, row.descricao_soma, clear=False)
                    vdoc = self._input_value(self.NUM_DOCUMENTO_MODAL) or row.descricao_soma
                    self._emit(f"Número do documento preenchido com sucesso: {vdoc}", row=row.row_number, tipo=row.tipo.value)
                except Exception:
                    pass

            chosen_cx = self._select_with_sleep_validation(
                self._resolve_caixa_pagamento_modal_select,
                row.caixa,
                row=row,
                field="CAIXA_PAGAMENTO_MODAL",
                strict=self.strict_caixa,
            )
            self._emit(f"Caixa para pagamento selecionada com sucesso: {chosen_cx}", row=row.row_number, tipo=row.tipo.value)

            self.a.click_js(self.BTN_SALVAR_PAGAMENTO_MODAL)
            time.sleep(2)
            self._emit("Botão 'Salvar Pagamento' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._dismiss_overlays()
            self._emit("Botão 'OK Baixa' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

            if self.strict_caixa and not self._match_ok(row.caixa, chosen_cx):
                raise RuntimeError(f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{row.caixa}' | selecionado='{chosen_cx}'")

    def _do_baixa(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.baixa", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._dismiss_overlays()
            self.a.click_js(self.BTN_INSERIR_BAIXA)
            time.sleep(1)
            self._emit("Inserir Baixa salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)

            self.a.type(self.DATA_BAIXA, row.data_mov)
            time.sleep(0.5)
            v = self._input_value(self.DATA_BAIXA) or row.data_mov
            self._emit(f"Data início preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

            try:
                loc = self.a.wait_any_present(self.POPUP_CLICK_CANDIDATES, timeout_seconds=3)
                self.a.click_js(loc)
                self._emit("Click na janela pop-up com sucesso", row=row.row_number, tipo=row.tipo.value)
            except Exception:
                pass

            self.a.click_js(self.BTN_SALVAR_BAIXA)
            time.sleep(2)
            self._emit("Salvar Baixa salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._dismiss_overlays()
            self._emit("Botão 'OK Baixa' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

    # -----------------------
    # doc search
    # -----------------------
    def _go_back_to_list_best_effort(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.back_to_list_best_effort", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays()
            self._close_datepicker()
            try:
                if self.a.exists(self.BTN_VOLTAR, timeout_seconds=2):
                    self.a.click_js(self.BTN_VOLTAR)
                    self.a.wait_dom_ready(15)
                    time.sleep(1)
            except Exception:
                pass
            self._ensure_pesquisa_visivel(row)

    def _search_doc_id(self, row: ContaOrdemRow) -> str:
        with step(log, "entradas_saidas.search_doc", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._go_back_to_list_best_effort(row)

            self.a.type(self.PESQ_DESCRICAO, row.descricao_soma)
            self._emit(f"Campo pesquisar a descrição preenchida com sucesso: {row.descricao_soma}", row=row.row_number, tipo=row.tipo.value)

            self.a.click_js(self.RADIO_PERIODO)
            time.sleep(0.5)
            self._emit("Selecionado o botão de rádio 'Periodo'", row=row.row_number, tipo=row.tipo.value)

            self.a.click_js(self.RADIO_DATA_PAGAMENTO)
            time.sleep(0.5)
            self._emit("Selecionado o botão de rádio 'Data de Pagamento'", row=row.row_number, tipo=row.tipo.value)

            self.a.type(self.DATA_INI, row.data_mov)
            self._emit(f"Data início preenchida com sucesso: {row.data_mov}", row=row.row_number, tipo=row.tipo.value)

            self.a.type(self.DATA_FIM, row.data_mov)
            self._emit(f"Data fim preenchida com sucesso: {row.data_mov}", row=row.row_number, tipo=row.tipo.value)

            self.a.click_js(self.BTN_PESQUISAR)
            self._emit("Botão 'Pesquisar' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

            try:
                doc = self.a.wait_visible(self.RESULT_DOC, timeout_seconds=30).text.strip()
            except TimeoutException as e:
                if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=2):
                    raise RuntimeError(
                        f"Sem resultados na pesquisa do nº SOMA. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
                    ) from e
                raise TimeoutException(
                    f"Timeout à espera do RESULT_DOC. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
                ) from e

            if not doc:
                raise RuntimeError("Doc ID vazio após pesquisa.")

            self._emit(f"Número do documento extraído: {doc}", row=row.row_number, tipo=row.tipo.value)
            return doc

    def fetch_dados_doc(self, doc_id: str) -> str:
        url = f"{self.base_ivv}?mod=ivv&exec=entradas_saidas_dados&ID={doc_id}"
        with step(log, "entradas_saidas.fetch_dados_doc", doc=doc_id, url=url):
            self._emit(f"Redirecionando para: {url}")
            self.a.driver.get(url)
            self.a.wait_dom_ready(20)
            self._emit("Nova página carregada com sucesso!")
            cell = self.a.wait_visible(self.DADOS_DOC_CELL, timeout_seconds=30)
            txt = (cell.text or "").strip()
            self._emit(f"Número do documento extraído: {txt}")
            return txt

    def recover_doc_id(self, row: ContaOrdemRow) -> str:
        return self._search_doc_id(row)

    def create_and_get_doc_id(self, row: ContaOrdemRow) -> str:
        with step(log, "entradas_saidas.create_start", row=row.row_number, tipo=row.tipo.value):
            self._open_new(row)
            self._choose_tipo(row)

            with step(log, "entradas_saidas.fill_form", row=row.row_number, tipo=row.tipo.value):
                self._fill_common(row)
                if row.tipo == TipoMovimento.SAIDA:
                    self._fill_saida(row)
                else:
                    self._fill_entrada_sem_caixa(row)
                    self._fill_caixa_entrada_ultima(row)

            self._save_form_if_present(row)

            if row.tipo == TipoMovimento.SAIDA:
                self._realizar_pagamento(row)
                self._pagamento_saida_modal(row)
                if (row.forma_pagamento or "").strip().upper() == "TRANSFERÊNCIA BANCÁRIA":
                    self._do_baixa(row)
            else:
                if (row.forma_pagamento or "").strip().upper() == "TRANSFERÊNCIA BANCÁRIA":
                    self._realizar_pagamento(row)
                    self._do_baixa(row)

            doc = self._search_doc_id(row)
            log_kv(log, logging.INFO, "Documento criado.", row=row.row_number, tipo=row.tipo.value, doc=doc)
            return doc