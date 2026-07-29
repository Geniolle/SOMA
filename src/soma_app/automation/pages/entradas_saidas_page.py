from __future__ import annotations

import logging
import time
import unicodedata
from datetime import datetime, timedelta
from typing import Any, Callable, Iterable, List, Tuple

from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from soma_app.automation.actions import Actions
from soma_app.config.locators import apply_locator_overrides
from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra.trace import log_kv, step

log = logging.getLogger("soma_app.pages.entradas_saidas")


class EntradasSaidasPage:
    # =========
    # MENU / NOVO (robusto por candidatos)
    # =========
    MENU_ENTRADAS_SAIDAS_CANDIDATES = []
    BTN_NOVA_CANDIDATES = []

    # =========
    # RADIO TIPO
    # =========
    RADIO_SAIDA = (By.XPATH, "")
    RADIO_ENTRADA = (By.XPATH, "")

    FORM_CONTAINER = (By.XPATH, "")

    RADIO_ENTRADA_CANDIDATES = []
    RADIO_SAIDA_CANDIDATES = []
    RADIO_ANY_CANDIDATES = []

    # =========
    # CAMPOS COMUNS
    # =========
    PLANO_CONTA = (By.XPATH, "")
    CENTRO_CUSTO = (By.XPATH, "")
    DESCRICAO = (By.XPATH, "")
    VALOR = (By.XPATH, "")
    OBS = (By.XPATH, "")

    PLANO_CONTA_CANDIDATES = []
    CENTRO_CUSTO_CANDIDATES = []
    DESCRICAO_CANDIDATES = []
    OBS_CANDIDATES = []

    # =========
    # CAMPOS ENTRADA
    # =========
    DATA_ENTRADA = (By.XPATH, "")
    FORMA_PAGAMENTO_ENTRADA = (By.XPATH, "")
    CAIXA_ENTRADA_CONTAINER = (By.XPATH, "")

    FORMA_PAGAMENTO_ENTRADA_CANDIDATES = []

    # =========
    # CAMPOS SAÍDA
    # =========
    DATA_VENCIMENTO_SAIDA = (By.XPATH, "")

    # =========
    # BOTÕES (form)
    # =========
    BTN_SALVAR_FORM = (By.XPATH, "")
    BTN_VOLTAR = (By.XPATH, "")

    FORM_READY_SELECT2 = (By.XPATH, "")
    FORM_READY_CANDIDATES = []

    # =========
    # ALERTAS / OVERLAY
    # =========
    OK_ALERT = (By.XPATH, "")
    SWAL_CONTAINER = (By.CLASS_NAME, "swal2-container")

    # =========
    # PAGAMENTO / BAIXA
    # =========
    BTN_REALIZAR_PAGAMENTO = (By.XPATH, "")
    BTN_REALIZAR_PAGAMENTO_CANDIDATES = []

    BTN_INSERIR_PAGAMENTO_SAIDA = (By.XPATH, "")
    BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = []
    DATA_PAGAMENTO_MODAL = (By.XPATH, "")
    FORMA_PAGAMENTO_MODAL = (By.XPATH, "")
    CAIXA_PAGAMENTO_MODAL = (By.XPATH, "")
    NUM_DOCUMENTO_MODAL = (By.XPATH, "")
    BTN_SALVAR_PAGAMENTO_MODAL = (By.XPATH, "")
    BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES = []

    BTN_INSERIR_BAIXA = (By.XPATH, "")
    BTN_INSERIR_BAIXA_CANDIDATES = []
    DATA_BAIXA = (By.XPATH, "")
    POPUP_CLICK_CANDIDATES = []
    BTN_SALVAR_BAIXA = (By.XPATH, "")
    BTN_SALVAR_BAIXA_CANDIDATES = []

    # =========
    # PESQUISA DOC SOMA
    # =========
    PESQ_DESCRICAO = (By.XPATH, "")
    RADIO_PERIODO = (By.XPATH, "")
    RADIO_DATA_PAGAMENTO = (By.XPATH, "")
    DATA_INI = (By.XPATH, "")
    DATA_FIM = (By.XPATH, "")
    BTN_PESQUISAR = (By.XPATH, "")
    RESULT_DOC = (By.XPATH, "")

    NO_RESULTS_CANDIDATES = []

    # =========
    # DADOS DOC
    # =========
    DADOS_DOC_CELL = (By.XPATH, "")
    DADOS_DOC_CANDIDATES = []

    def __init__(self, actions: Actions, settings: Any):
        self.a = actions
        self.settings = settings
        self.base_ivv = (getattr(settings, "site_home_url", "") or "https://verbodavida.info/IVV/").rstrip("/") + "/"
        apply_locator_overrides(self, "entradas_saidas")
        self.RADIO_ANY_CANDIDATES = self.RADIO_SAIDA_CANDIDATES + self.RADIO_ENTRADA_CANDIDATES
        self.FORM_READY_CANDIDATES = [
            self.FORM_CONTAINER,
            self.DESCRICAO,
            self.BTN_SALVAR_FORM,
            self.FORM_READY_SELECT2,
        ]

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
            log_kv(log, msg, level=level, **kv)
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

    @staticmethod
    def _unique_locators(locators: Iterable[Tuple[str, str]]) -> List[Tuple[str, str]]:
        seen: set[Tuple[str, str]] = set()
        out: List[Tuple[str, str]] = []
        for loc in locators:
            if loc in seen:
                continue
            seen.add(loc)
            out.append(loc)
        return out

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

    def _select2_choose_candidates(
        self,
        locators: Iterable[Tuple[str, str]],
        value: str,
        *,
        row: ContaOrdemRow,
        field: str,
        wait_seconds: int = 20,
    ) -> None:
        unique = self._unique_locators(locators)
        last_error: Exception | None = None

        try:
            self.a.wait_any_present(unique, timeout_seconds=wait_seconds)
        except Exception as e:
            p = self.a.screenshot(f"entradas_saidas_{field.lower()}_opener_missing_row_{row.row_number}")
            raise TimeoutException(
                f"{field}: opener do Select2 não apareceu | value='{self._safe(value)}' | screenshot={p}"
            ) from e

        for loc in unique:
            if not self.a.exists(loc, timeout_seconds=2):
                continue
            try:
                self.a.select2_choose(loc, value)
                return
            except Exception as e:
                last_error = e

        p = self.a.screenshot(f"entradas_saidas_{field.lower()}_select2_fail_row_{row.row_number}")
        raise RuntimeError(
            f"{field}: não consegui selecionar '{self._safe(value)}' no Select2. "
            f"last_err={last_error} | screenshot={p}"
        )

    def _type_and_validate_candidates(
        self,
        locators: Iterable[Tuple[str, str]],
        value: str,
        *,
        row: ContaOrdemRow,
        field: str,
        clear: bool = True,
        click_first: bool = False,
        wait_seconds: int = 20,
    ) -> str:
        unique = self._unique_locators(locators)
        last_error: Exception | None = None
        last_seen = ""

        try:
            self.a.wait_any_present(unique, timeout_seconds=wait_seconds)
        except Exception as e:
            p = self.a.screenshot(f"entradas_saidas_{field.lower()}_missing_row_{row.row_number}")
            raise TimeoutException(
                f"{field}: campo não apareceu | value='{self._safe(value)}' | screenshot={p}"
            ) from e

        for loc in unique:
            if not self.a.exists(loc, timeout_seconds=2):
                continue
            for _ in range(2):
                try:
                    if click_first:
                        try:
                            self.a.click_js(loc)
                        except Exception:
                            pass
                    self.a.type(loc, value, clear=clear)
                    last_seen = self._input_value(loc)
                    if self._match_ok(value, last_seen):
                        return last_seen
                    time.sleep(0.2)
                except Exception as e:
                    last_error = e
                    break

        p = self.a.screenshot(f"entradas_saidas_{field.lower()}_fill_fail_row_{row.row_number}")
        raise RuntimeError(
            f"{field}: campo não confirmou preenchimento. esperado='{self._safe(value)}' | "
            f"atual='{self._safe(last_seen)}' | screenshot={p} | last_err={last_error}"
        )

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
                prev = self._select_first_text(sel_el)
                for _ in range(max(1, self.caixa_stable_checks)):
                    try:
                        cur = self._select_first_text(sel_el)
                    except StaleElementReferenceException:
                        stable = False
                        break
                    if cur != prev:
                        stable = False
                        break
                    time.sleep(self.caixa_stable_interval)

                try:
                    current = self._select_first_text(sel_el)
                except StaleElementReferenceException:
                    current = ""

                last_seen = current or chosen_now

                if stable and self._match_ok(desired, last_seen):
                    return last_seen

                try:
                    self._dump_caixa_diagnostics(
                        row,
                        suffix=f"attempt_{attempt}_unstable",
                        desired=desired,
                        last_seen=last_seen,
                    )
                except Exception:
                    pass

                self._emit(
                    f"Aviso: {field} alterou após seleção. tentativa {attempt}/{self.caixa_validate_retries} | "
                    f"esperado='{desired}' | agora='{last_seen}'. Reaplicando...",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )

            except StaleElementReferenceException:
                try:
                    self._dump_caixa_diagnostics(
                        row,
                        suffix=f"attempt_{attempt}_stale",
                        desired=desired,
                        last_seen=last_seen,
                    )
                except Exception:
                    pass
                self._emit(
                    f"Aviso: {field} ficou stale durante a seleção. tentativa {attempt}/{self.caixa_validate_retries}. Reaplicando...",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                time.sleep(0.3)
                continue

        if strict:
            try:
                self._dump_caixa_diagnostics(
                    row,
                    suffix="final_failed",
                    desired=desired,
                    last_seen=last_seen,
                )
            except Exception:
                pass
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
        self._dismiss_overlays_with_wait(max_wait_seconds=3)

    def _dismiss_overlays_with_wait(self, max_wait_seconds: int = 3) -> None:
        clicked = False
        for _ in range(2):
            try:
                if self.a.exists(self.OK_ALERT, timeout_seconds=1):
                    self.a.click_js(self.OK_ALERT)
                    self._emit("Botão 'OK' clicado com sucesso!")
                    clicked = True
                    time.sleep(0.2)
                    break
            except Exception:
                pass

        try:
            if self.a.exists(self.SWAL_CONTAINER, timeout_seconds=1):
                if not clicked:
                    try:
                        self.a.driver.switch_to.active_element.send_keys(Keys.ESCAPE)
                        time.sleep(0.2)
                    except Exception:
                        pass

                try:
                    self.a.wait_invisible(self.SWAL_CONTAINER, timeout_seconds=max_wait_seconds)
                except Exception:
                    self._emit(
                        "Aviso: O modal SweetAlert ainda está visível após o tempo de espera.",
                        level=logging.WARNING,
                    )
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
                        "Form 'Nova Entrada/Saída' não ficou pronto. Vou tentar novamente.",
                        level=logging.ERROR,
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
            self.a.wait_any_present(self.FORM_READY_CANDIDATES, timeout_seconds=30)
            time.sleep(1)

    # -----------------------
    # preenchimento
    # -----------------------
    def _fill_common(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.plano_conta", row=row.row_number, tipo=row.tipo.value, field="PLANO_CONTA"):
            self._select2_choose_candidates(
                self.PLANO_CONTA_CANDIDATES,
                row.plano_conta,
                row=row,
                field="plano_conta",
            )
            self._emit(f"Plano de conta preenchido com sucesso: {row.plano_conta}", row=row.row_number, tipo=row.tipo.value)

        with step(log, "entradas_saidas.fill.centro_custo", row=row.row_number, tipo=row.tipo.value, field="CENTRO_CUSTO"):
            self._select2_choose_candidates(
                self.CENTRO_CUSTO_CANDIDATES,
                row.centro_custo,
                row=row,
                field="centro_custo",
            )
            self._emit(f"Centro de custo preenchido com sucesso: {row.centro_custo}", row=row.row_number, tipo=row.tipo.value)

        with step(log, "entradas_saidas.fill.descricao", row=row.row_number, tipo=row.tipo.value, field="DESCRICAO"):
            v = self._type_and_validate_candidates(
                self.DESCRICAO_CANDIDATES,
                row.descricao_soma,
                row=row,
                field="descricao",
                clear=True,
            )
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
            v = self._type_and_validate_candidates(
                self.OBS_CANDIDATES,
                row.descricao_soma,
                row=row,
                field="obs",
                clear=True,
                click_first=True,
            )
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
            self._select2_choose_candidates(
                self.FORMA_PAGAMENTO_ENTRADA_CANDIDATES,
                row.forma_pagamento,
                row=row,
                field="forma_pagamento",
            )
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

                try:
                    chosen = self._select_with_sleep_validation(
                        self._resolve_caixa_entrada_select,
                        row.caixa,
                        row=row,
                        field="CAIXA_ENTRADA",
                        strict=self.strict_caixa,
                    )
                except Exception:
                    try:
                        self._dump_caixa_diagnostics(
                            row,
                            suffix="fill_exception",
                            desired=row.caixa,
                        )
                    except Exception:
                        pass
                    raise
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
                self._dismiss_overlays_with_wait(max_wait_seconds=3)
                self._close_datepicker()
                if self.a.exists(self.BTN_SALVAR_FORM, timeout_seconds=2):
                    self.a.click_js(self.BTN_SALVAR_FORM)
                    time.sleep(1)
                    self._dismiss_overlays_with_wait(max_wait_seconds=3)
                    self._emit("Campos principais preenchidos com sucesso")
            except Exception:
                pass

    # -----------------------
    # pagamento/baixa
    # -----------------------
    def _realizar_pagamento(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.realizar_pagamento", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays_with_wait(max_wait_seconds=3)
            candidates = self._unique_locators(self.BTN_REALIZAR_PAGAMENTO_CANDIDATES or [self.BTN_REALIZAR_PAGAMENTO])
            if not self._exists_any(candidates, timeout_seconds=2):
                self._emit(
                    "Botão de realizar pagamento não apareceu no DOM; vou seguir para o modal/continuação do fluxo.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            loc = self.a.wait_any_present(candidates, timeout_seconds=10)
            self.a.click_js(loc)
            time.sleep(1)
            self._dismiss_overlays_with_wait(max_wait_seconds=3)
            self._emit("Realizar pagamento salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)

    def _resolve_caixa_pagamento_modal_select(self) -> Any:
        return self.a.driver.find_element(*self.CAIXA_PAGAMENTO_MODAL)

    def _pagamento_saida_modal(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.pagamento_saida_modal", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays_with_wait(max_wait_seconds=3)

            candidates = self._unique_locators(self.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES or [self.BTN_INSERIR_PAGAMENTO_SAIDA])
            if not self._exists_any(candidates, timeout_seconds=2):
                self._emit(
                    "Botão de inserir pagamento não apareceu no DOM; vou seguir sem abrir o modal dedicado.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            loc = self.a.wait_any_present(candidates, timeout_seconds=10)
            self.a.click_js(loc)
            time.sleep(0.8)
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
                    vdoc = self._type_and_validate_candidates(
                        [self.NUM_DOCUMENTO_MODAL],
                        row.descricao_soma,
                        row=row,
                        field="num_documento_modal",
                        clear=True,
                    )
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

            salvar_candidates = self._unique_locators(
                self.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES or [self.BTN_SALVAR_PAGAMENTO_MODAL]
            )
            salvar_loc = self.a.wait_any_present(salvar_candidates, timeout_seconds=10)
            self.a.click_js(salvar_loc)
            time.sleep(1)
            self._emit("Botão 'Salvar Pagamento' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._dismiss_overlays_with_wait(max_wait_seconds=3)
            self._emit("Botão 'OK Baixa' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

            if self.strict_caixa and not self._match_ok(row.caixa, chosen_cx):
                raise RuntimeError(f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{row.caixa}' | selecionado='{chosen_cx}'")

    def _do_baixa(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.baixa", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._dismiss_overlays_with_wait(max_wait_seconds=3)
            candidates = self._unique_locators(self.BTN_INSERIR_BAIXA_CANDIDATES or [self.BTN_INSERIR_BAIXA])
            if not self._exists_any(candidates, timeout_seconds=2):
                self._emit(
                    "Botão de inserir baixa não apareceu no DOM; vou seguir sem abrir a baixa dedicada.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            loc = self.a.wait_any_present(candidates, timeout_seconds=10)
            self.a.click_js(loc)
            time.sleep(0.8)
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

            salvar_candidates = self._unique_locators(self.BTN_SALVAR_BAIXA_CANDIDATES or [self.BTN_SALVAR_BAIXA])
            salvar_loc = self.a.wait_any_present(salvar_candidates, timeout_seconds=10)
            self.a.click_js(salvar_loc)
            time.sleep(1)
            self._emit("Salvar Baixa salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._dismiss_overlays_with_wait(max_wait_seconds=3)
            self._emit("Botão 'OK Baixa' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

    # -----------------------
    # doc search
    # -----------------------
    def _click_radio_force_change(self, locator: Tuple[str, str]) -> None:
        # a revelação do painel de data é um TOGGLE ligado ao clique do rádio,
        # não um "show" condicional: clicar de novo num retry (mesma página,
        # sem reload, painel já visível da tentativa anterior) esconde-o outra
        # vez. Por isso só clicamos se o rádio ainda não estiver marcado.
        el = self.a.wait_present(locator, timeout_seconds=30)
        if el.is_selected():
            return
        self.a.click_js(locator)

    def _go_back_to_list_best_effort(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.back_to_list_best_effort", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays()
            self._close_datepicker()
            if self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=2):
                return

            used_fallback = False
            try:
                if self.a.exists(self.BTN_VOLTAR, timeout_seconds=2):
                    self.a.click_js(self.BTN_VOLTAR)
                    self.a.wait_dom_ready(15)
                    time.sleep(0.5)
                    used_fallback = True
            except Exception:
                pass

            if not self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=2):
                try:
                    self.a.driver.back()
                    self.a.wait_dom_ready(15)
                    time.sleep(0.5)
                    used_fallback = True
                except Exception:
                    pass

            if not self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=2):
                try:
                    self._dump_search_diagnostics(row, suffix="back_failed")
                except Exception:
                    pass
                if used_fallback:
                    self._emit(
                        "Retorno para a lista exigiu fallback adicional; vou reabrir o painel de pesquisa.",
                        level=logging.WARNING,
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
            self._ensure_pesquisa_visivel(row)

    def _search_doc_id_attempt(self, row: ContaOrdemRow) -> str:
        self._go_back_to_list_best_effort(row)

        self.a.type(self.PESQ_DESCRICAO, row.descricao_soma)
        self._emit(f"Campo pesquisar a descrição preenchida com sucesso: {row.descricao_soma}", row=row.row_number, tipo=row.tipo.value)

        self._click_radio_force_change(self.RADIO_PERIODO)
        time.sleep(0.5)
        self._emit("Selecionado o botão de rádio 'Periodo'", row=row.row_number, tipo=row.tipo.value)

        self._click_radio_force_change(self.RADIO_DATA_PAGAMENTO)
        time.sleep(0.5)
        self._emit("Selecionado o botão de rádio 'Data de Pagamento'", row=row.row_number, tipo=row.tipo.value)

        if not self.a.exists(self.DATA_INI, timeout_seconds=3):
            # painel pode ter ficado escondido mesmo com o rádio já marcado
            # (ex.: reset parcial ao voltar à lista); força mais um toggle.
            self.a.click_js(self.RADIO_DATA_PAGAMENTO)
            time.sleep(0.5)

        self.a.type(self.DATA_INI, row.data_mov)
        self._emit(f"Data início preenchida com sucesso: {row.data_mov}", row=row.row_number, tipo=row.tipo.value)

        self.a.type(self.DATA_FIM, row.data_mov)
        self._emit(f"Data fim preenchida com sucesso: {row.data_mov}", row=row.row_number, tipo=row.tipo.value)

        self.a.click_js(self.BTN_PESQUISAR)
        self._emit("Botão 'Pesquisar' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

        try:
            doc = self._read_search_result_doc(timeout_seconds=15)
        except TimeoutException as e:
            try:
                self._dump_search_diagnostics(row, suffix="timeout")
            except Exception:
                pass
            if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=2):
                raise RuntimeError(
                    f"Sem resultados na pesquisa do nº SOMA. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
                ) from e
            raise TimeoutException(
                f"Timeout à espera do RESULT_DOC. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
            ) from e

        if not doc:
            try:
                self._dump_search_diagnostics(row, suffix="empty")
            except Exception:
                pass
            raise RuntimeError("Doc ID vazio após pesquisa.")

        self._emit(f"Número do documento extraído: {doc}", row=row.row_number, tipo=row.tipo.value)
        return doc

    def _dump_search_diagnostics(self, row: ContaOrdemRow, *, suffix: str) -> None:
        name = f"entradas_saidas_search_doc_row_{row.row_number}_{suffix}"
        self.a.dump_page_source(name)
        self.a.dump_locator_probe(name, [self.PESQ_DESCRICAO, self.DATA_INI, self.DATA_FIM, self.BTN_PESQUISAR, self.RESULT_DOC])

    def _dump_caixa_diagnostics(self, row: ContaOrdemRow, *, suffix: str, desired: str, last_seen: str = "") -> None:
        name = f"entradas_saidas_caixa_row_{row.row_number}_{suffix}"
        self.a.screenshot(name)
        self.a.dump_page_source(name)
        self.a.dump_locator_probe(name, [self.CAIXA_ENTRADA_CONTAINER])
        log_kv(
            log,
            "Diagnostico de CAIXA gravado.",
            level=logging.WARNING,
            row=row.row_number,
            tipo=row.tipo.value,
            desired=desired,
            last_seen=last_seen,
            artifact=name,
        )

    def fetch_dados_doc(self, doc_id: str) -> str:
        url = f"{self.base_ivv}index.php?mod=ivv&exec=entradas_saidas"
        with step(log, "entradas_saidas.fetch_dados_doc", doc=doc_id, url=url):
            self._emit(f"Redirecionando para: {url}")
            self.a.driver.get(url)
            self.a.wait_dom_ready(20)
            self._emit("Nova página carregada com sucesso!")
            candidates = self._unique_locators(self.DADOS_DOC_CANDIDATES or ([self.DADOS_DOC_CELL] if self.DADOS_DOC_CELL[1].strip() else []))
            end = time.time() + 12
            last_seen = ""
            while time.time() < end:
                for loc in candidates:
                    try:
                        if self.a.exists(loc, timeout_seconds=0):
                            cell = self.a.driver.find_element(*loc)
                            txt = (cell.text or "").strip()
                            if txt.isdigit():
                                self._emit(f"Número do documento extraído: {txt}")
                                return txt
                    except Exception:
                        pass

                try:
                    cells = self.a.driver.find_elements(By.CSS_SELECTOR, "table tbody tr td")
                    for cell in cells:
                        txt = (cell.text or "").strip()
                        if not txt:
                            continue
                        last_seen = txt
                        if txt.isdigit():
                            self._emit(f"Número do documento extraído: {txt}")
                            return txt
                except Exception:
                    pass

                time.sleep(0.25)

            self._emit(
                f"Falha ao extrair dados do documento; vou devolver o doc_id original. last_seen='{self._safe(last_seen)}'",
                level=logging.WARNING,
                doc=doc_id,
            )
            return doc_id

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
            log_kv(log, "Documento criado.", level=logging.INFO, row=row.row_number, tipo=row.tipo.value, doc=doc)
            return doc

    def _read_search_result_doc(self, timeout_seconds: int = 15) -> str:
        end = time.time() + timeout_seconds
        while time.time() < end:
            try:
                if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=0):
                    return ""
            except Exception:
                pass

            try:
                if self.a.exists(self.RESULT_DOC, timeout_seconds=0):
                    doc = (self.a.driver.find_element(*self.RESULT_DOC).text or "").strip()
                    if doc.isdigit():
                        return doc
            except Exception:
                pass

            try:
                rows = self.a.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                for tr in rows:
                    cells = tr.find_elements(By.CSS_SELECTOR, "td,th")
                    if not cells:
                        continue
                    for cell in cells:
                        txt = (cell.text or "").strip()
                        if not txt:
                            continue
                        if txt.isdigit():
                            return txt
            except Exception:
                pass

            time.sleep(0.3)

        raise TimeoutException("Timeout à espera do resultado da pesquisa do nº SOMA.")

    def _search_doc_id(self, row: ContaOrdemRow) -> str:
        with step(log, "entradas_saidas.search_doc", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            attempts = 2
            delays = (2,)
            last_exc: Exception | None = None
            for attempt in range(1, attempts + 1):
                try:
                    return self._search_doc_id_attempt(row)
                except RuntimeError as e:
                    last_exc = e
                    err_text = str(e)
                    if ("Sem resultados" not in err_text and "doc vazio" not in err_text.lower()) or attempt == attempts:
                        break
                    delay = delays[min(attempt - 1, len(delays) - 1)]
                    self._emit(
                        f"Pesquisa sem resultados (tentativa {attempt}/{attempts}), retry em {delay}s",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    time.sleep(delay)

            try:
                return self._search_doc_id_broader(row)
            except Exception as broader_exc:
                last_exc = broader_exc

            assert last_exc is not None
            raise last_exc

    @staticmethod
    def _date_window_variants(date_text: str) -> List[tuple[str, str]]:
        try:
            base = datetime.strptime((date_text or "").strip(), "%d/%m/%Y")
        except Exception:
            return []

        windows = [
            (base, base),
            (base - timedelta(days=2), base + timedelta(days=2)),
            (base - timedelta(days=7), base + timedelta(days=7)),
        ]

        out: list[tuple[str, str]] = []
        seen: set[tuple[str, str]] = set()
        for start, end in windows:
            pair = (start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y"))
            if pair in seen:
                continue
            seen.add(pair)
            out.append(pair)
        return out

    def _search_doc_id_broader(self, row: ContaOrdemRow) -> str:
        windows = self._date_window_variants(row.data_mov)
        if not windows:
            raise RuntimeError("Não consegui gerar janelas de data para recovery do DOC.")

        last_exc: Exception | None = None
        for start_date, end_date in windows:
            try:
                self._go_back_to_list_best_effort(row)
                self.a.type(self.PESQ_DESCRICAO, row.descricao_soma)
                self._click_radio_force_change(self.RADIO_PERIODO)
                time.sleep(0.2)
                self._click_radio_force_change(self.RADIO_DATA_PAGAMENTO)
                time.sleep(0.2)
                self.a.type(self.DATA_INI, start_date)
                self.a.type(self.DATA_FIM, end_date)
                self.a.click_js(self.BTN_PESQUISAR)
                self._emit(
                    f"Recovery search com janela alargada: {start_date} -> {end_date}",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                doc = self._read_search_result_doc(timeout_seconds=12)
                if doc:
                    self._emit(
                        f"Número do documento extraído por recovery alargado: {doc}",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    return doc
            except Exception as e:
                last_exc = e
                continue

        try:
            self._dump_search_diagnostics(row, suffix="broader_failed")
        except Exception:
            pass
        if last_exc:
            raise last_exc
        raise RuntimeError("Recovery alargado falhou sem exceção explícita.")
