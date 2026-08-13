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
from soma_app.automation.debug_session import GuidedDebugSession
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
    PLANO_CONTA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span/span[1]")
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

    BTN_INSERIR_PAGAMENTO_SAIDA = (By.XPATH, "/html/body/div[2]/div/div[2]/div/div/form/div[29]/div/div[2]/a")
    DATA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input")
    FORMA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select")
    CAIXA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[4]/div[2]/div/select")
    NUM_DOCUMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[3]/div/input")
    BTN_SALVAR_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[3]/button")

    BTN_INSERIR_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button")
    BTN_INSERIR_BAIXA_CANDIDATES = [
        (By.XPATH, "//button[@id and contains(@data-target, 'inserirBaixa')]"),
        (By.XPATH, "//button[contains(., 'Inserir Baixa')]"),
        (By.XPATH, "//a[contains(@data-target, 'inserir') and contains(., 'Inserir')]"),
        (By.CSS_SELECTOR, "button[data-toggle='modal']"),
    ]

    # Baixa - campo de data dentro do modal #inserirBaixa (visível)
    DATA_BAIXA_CANDIDATES = [
        (By.XPATH, "//*[@id='inserirBaixa']//input[@name='data_baixa']"),
        (By.XPATH, "//*[@id='inserirBaixa']//input[@class and contains(@class, 'datepicker')]"),
        (By.CSS_SELECTOR, "#inserirBaixa input[name='data_baixa']"),
        (By.CSS_SELECTOR, "#inserirBaixa input.datepicker"),
    ]

    # Baixa - forma de pagamento dentro do modal #inserirBaixa
    FORMA_PAGAMENTO_BAIXA_CANDIDATES = [
        (By.XPATH, "//*[@id='inserirBaixa']//select[@name='forma_pagamento']"),
        (By.CSS_SELECTOR, "#inserirBaixa select[name='forma_pagamento']"),
    ]

    # Baixa - numero de documento dentro do modal #inserirBaixa
    NUM_DOCUMENTO_BAIXA_CANDIDATES = [
        (By.XPATH, "//*[@id='inserirBaixa']//input[@name='num_documento']"),
        (By.CSS_SELECTOR, "#inserirBaixa input[name='num_documento']"),
    ]

    # Baixa - caixa dentro do modal #inserirBaixa
    CAIXA_BAIXA_CANDIDATES = [
        (By.XPATH, "//*[@id='inserirBaixa']//select[@name='caixa' or contains(@class, 'caixa')]"),
        (By.CSS_SELECTOR, "#inserirBaixa select[name='caixa']"),
    ]

    # Baixa - botão salvar dentro do modal #inserirBaixa
    BTN_SALVAR_BAIXA_CANDIDATES = [
        (By.XPATH, "//*[@id='inserirBaixa']//button[contains(., 'Salvar') or contains(., 'OK')]"),
        (By.CSS_SELECTOR, "#inserirBaixa button[type='submit']"),
        (By.CSS_SELECTOR, "#inserirBaixa .modal-footer button"),
    ]

    POPUP_CLICK_CANDIDATES = []

    # Fallbacks absolutos (não use como primeiro candidato)
    DATA_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input")
    BTN_SALVAR_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[4]/div/div/form/div[3]/button")

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

    def __init__(self, actions: Actions, settings: Any):
        self.a = actions
        self.settings = settings
        self.debug_session = GuidedDebugSession(actions, settings)
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

    def _debug_checkpoint(
        self,
        *,
        row: ContaOrdemRow,
        stage: str,
        phase: str,
        action: str,
        element_name: str | None = None,
        locator: Tuple[str, str] | None = None,
        value: Any = None,
        instructions: Iterable[str] | None = None,
    ) -> None:
        if self.debug_session is None:
            return
        self.debug_session.checkpoint(
            row=row,
            stage=stage,
            phase=phase,
            action=action,
            element_name=element_name,
            locator=locator,
            value=value,
            instructions=instructions,
        )

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

    def _raise_fixed_xpath_error(
        self,
        *,
        row: ContaOrdemRow,
        stage: str,
        element_name: str,
        locator: Tuple[str, str],
        error: Exception | str,
        cause: str | None = None,
    ) -> None:
        screenshot = self.a.screenshot(f"saida_xpath_error_row_{row.row_number}_{element_name.lower()}")
        html = self.a.dump_page_source(f"saida_xpath_error_row_{row.row_number}_{element_name.lower()}")
        url = getattr(self.a.driver, "current_url", "")
        causa = cause or (type(error).__name__ if isinstance(error, Exception) else type(error).__name__)
        msg = (
            f"[SAIDA][XPATH_ERROR] linha={row.row_number} elemento={element_name} "
            f"xpath={locator[1]} etapa={stage} url={url} causa={causa} erro={error} screenshot={screenshot} html={html}"
        )
        log.error(msg)
        if isinstance(error, Exception):
            raise TimeoutException(msg) from error
        raise TimeoutException(msg)

    def _wait_fixed_visible(
        self,
        locator: Tuple[str, str],
        *,
        row: ContaOrdemRow,
        stage: str,
        element_name: str,
        timeout_seconds: int = 20,
    ) -> Any:
        try:
            return self.a.wait_visible(locator, timeout_seconds=timeout_seconds)
        except Exception as exc:
            self._raise_fixed_xpath_error(
                row=row,
                stage=stage,
                element_name=element_name,
                locator=locator,
                error=exc,
            )

    def _wait_fixed_displayed(
        self,
        locator: Tuple[str, str],
        *,
        row: ContaOrdemRow,
        stage: str,
        element_name: str,
        timeout_seconds: int = 30,
        poll_seconds: float = 0.25,
    ) -> Any:
        deadline = time.time() + max(1, timeout_seconds)
        last_error: Exception | None = None
        while time.time() < deadline:
            try:
                elements = self.a.driver.find_elements(*locator)
            except Exception as exc:
                last_error = exc
                time.sleep(poll_seconds)
                continue

            for el in elements:
                try:
                    if el.is_displayed():
                        return el
                except StaleElementReferenceException as exc:
                    last_error = exc
                    continue
                except Exception as exc:
                    last_error = exc
                    continue

            time.sleep(poll_seconds)

        self._raise_fixed_xpath_error(
            row=row,
            stage=stage,
            element_name=element_name,
            locator=locator,
            error=TimeoutException(
                f"{element_name}: elemento não ficou visível dentro de {timeout_seconds}s"
            ),
            cause=type(last_error).__name__ if last_error is not None else "OK_ALERT_NOT_FOUND",
        )

    def _click_fixed_visible(
        self,
        locator: Tuple[str, str],
        *,
        row: ContaOrdemRow,
        stage: str,
        element_name: str,
        timeout_seconds: int = 20,
    ) -> None:
        el = self._wait_fixed_visible(
            locator,
            row=row,
            stage=stage,
            element_name=element_name,
            timeout_seconds=timeout_seconds,
        )
        try:
            self.a.driver.execute_script("arguments[0].click();", el)
        except Exception as exc:
            self._raise_fixed_xpath_error(
                row=row,
                stage=stage,
                element_name=element_name,
                locator=locator,
                error=exc,
            )

    def _type_fixed_visible(
        self,
        locator: Tuple[str, str],
        value: str,
        *,
        row: ContaOrdemRow,
        stage: str,
        element_name: str,
        timeout_seconds: int = 20,
        clear: bool = True,
    ) -> str:
        el = self._wait_fixed_visible(
            locator,
            row=row,
            stage=stage,
            element_name=element_name,
            timeout_seconds=timeout_seconds,
        )
        try:
            if clear:
                el.clear()
            el.send_keys(value)
        except Exception as exc:
            self._raise_fixed_xpath_error(
                row=row,
                stage=stage,
                element_name=element_name,
                locator=locator,
                error=exc,
            )
        try:
            return (el.get_attribute("value") or "").strip() or (value or "").strip()
        except Exception:
            return (value or "").strip()

    def _select_fixed_visible_text(
        self,
        locator: Tuple[str, str],
        value: str,
        *,
        row: ContaOrdemRow,
        stage: str,
        element_name: str,
        timeout_seconds: int = 20,
        stale_retries: int = 0,
        stale_label: str | None = None,
    ) -> str:
        desired = (value or "").strip()
        retries = max(1, int(stale_retries) or 1)
        label = stale_label or element_name
        last_error: Exception | None = None
        last_seen = ""

        for attempt in range(1, retries + 1):
            el = self._wait_fixed_visible(
                locator,
                row=row,
                stage=stage,
                element_name=element_name,
                timeout_seconds=timeout_seconds,
            )
            try:
                Select(el).select_by_visible_text(value)
            except StaleElementReferenceException as exc:
                last_error = exc
                if attempt < retries:
                    self._emit(
                        f"[SAIDA] {label} ficou stale. Relocalizando pelo mesmo XPath tentativa {attempt}/{stale_retries}",
                        level=logging.WARNING,
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    time.sleep(0.5)
                    continue
                break
            try:
                last_seen = (Select(el).first_selected_option.text or "").strip() or desired
            except StaleElementReferenceException as exc:
                last_error = exc
                if attempt < retries:
                    self._emit(
                        f"[SAIDA] {label} ficou stale. Relocalizando pelo mesmo XPath tentativa {attempt}/{stale_retries}",
                        level=logging.WARNING,
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    time.sleep(0.5)
                    continue
                break

            if desired and not self._match_ok(desired, last_seen):
                raise RuntimeError(
                    f"{element_name}: seleção não confirmou valor. esperado='{desired}' | atual='{last_seen}'"
                )

            if attempt > 1:
                self._emit(
                    f"[SAIDA] {label} localizada novamente",
                    level=logging.INFO,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
            return last_seen

        if last_error is None:
            last_error = StaleElementReferenceException(
                f"{element_name}: stale após {stale_retries} tentativas"
            )
        self._raise_fixed_xpath_error(
            row=row,
            stage=stage,
            element_name=element_name,
            locator=locator,
            error=last_error,
        )

    def _numero_documento_para_pagamento_saida(self, row: ContaOrdemRow) -> str:
        forma = self._norm(row.forma_pagamento)
        if forma == self._norm("TRANSFERÊNCIA BANCÁRIA"):
            numero_documento = (row.id_interno or "").strip()
            if not numero_documento:
                raise ValueError(f"Transferência Bancária requer ID_INTERNO preenchido. Linha {row.row_number}")
            return numero_documento
        return ""

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
                    self.a.wait_invisible(self.SWAL_CONTAINER, timeout_seconds=2)
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
    def _fill_saida_plano_conta(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.saida.plano_conta", row=row.row_number, tipo=row.tipo.value, field="PLANO_CONTA"):
            time.sleep(1.5)
            try:
                self._select2_choose_candidates(
                    self.PLANO_CONTA_CANDIDATES,
                    row.plano_conta,
                    row=row,
                    field="plano_conta_saida",
                    wait_seconds=25,
                )
            except Exception as e:
                p = self.a.screenshot(f"entradas_saidas_plano_conta_saida_fail_row_{row.row_number}")
                log_kv(log, "Falha ao preencher plano de conta para Saída", level=logging.ERROR,
                       row=row.row_number, erro=str(e), screenshot=p)
                raise
            self._emit(f"Plano de conta preenchido com sucesso (Saída): {row.plano_conta}", row=row.row_number, tipo=row.tipo.value)

    def _fill_common(self, row: ContaOrdemRow) -> None:
        # Ativa debug interativo para preenchimento de dados
        self.a.set_debug_context("input_dados")

        if row.tipo == TipoMovimento.SAIDA:
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.FORM.PLANO_CONTA",
                phase="BEFORE",
                action="SELECT",
                element_name="PLANO_CONTA",
                locator=self.PLANO_CONTA,
                value=row.plano_conta,
                instructions=[
                    "Confirme que o browser está visível e que o formulário de Saída está aberto.",
                    "Verifique visualmente o campo Plano de Conta.",
                    f"Se quiser testar o XPath, use: x {self.PLANO_CONTA[1]}",
                    "Pressione ENTER para o Selenium selecionar o Plano de Conta.",
                ],
            )
        with step(log, "entradas_saidas.fill.plano_conta", row=row.row_number, tipo=row.tipo.value, field="PLANO_CONTA"):
            self._select2_choose_candidates(
                self.PLANO_CONTA_CANDIDATES,
                row.plano_conta,
                row=row,
                field="plano_conta",
            )
            self._emit(f"Plano de conta preenchido com sucesso: {row.plano_conta}", row=row.row_number, tipo=row.tipo.value)
        if row.tipo == TipoMovimento.SAIDA:
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.FORM.PLANO_CONTA",
                phase="AFTER",
                action="SELECT",
                element_name="PLANO_CONTA",
                locator=self.PLANO_CONTA,
                value=row.plano_conta,
                instructions=[
                    "Confirme que o Plano de Conta ficou selecionado corretamente.",
                    "O próximo passo será selecionar o Centro de Custo.",
                    "Pressione ENTER para continuar.",
                ],
            )

        if row.tipo == TipoMovimento.SAIDA:
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.FORM.CENTRO_CUSTO",
                phase="BEFORE",
                action="SELECT",
                element_name="CENTRO_CUSTO",
                locator=self.CENTRO_CUSTO,
                value=row.centro_custo,
                instructions=[
                    "Confirme que o campo Centro de Custo está visível.",
                    f"Se quiser validar o XPath, use: x {self.CENTRO_CUSTO[1]}",
                    "Pressione ENTER para o Selenium selecionar o Centro de Custo.",
                ],
            )
        with step(log, "entradas_saidas.fill.centro_custo", row=row.row_number, tipo=row.tipo.value, field="CENTRO_CUSTO"):
            self._select2_choose_candidates(
                self.CENTRO_CUSTO_CANDIDATES,
                row.centro_custo,
                row=row,
                field="centro_custo",
            )
            self._emit(f"Centro de custo preenchido com sucesso: {row.centro_custo}", row=row.row_number, tipo=row.tipo.value)
        if row.tipo == TipoMovimento.SAIDA:
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.FORM.CENTRO_CUSTO",
                phase="AFTER",
                action="SELECT",
                element_name="CENTRO_CUSTO",
                locator=self.CENTRO_CUSTO,
                value=row.centro_custo,
                instructions=[
                    "Confirme que o Centro de Custo foi selecionado corretamente.",
                    "O próximo passo será preencher a Descrição.",
                    "Pressione ENTER para continuar.",
                ],
            )

        if row.tipo == TipoMovimento.SAIDA:
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.FORM.DESCRICAO",
                phase="BEFORE",
                action="INPUT",
                element_name="DESCRICAO",
                locator=self.DESCRICAO,
                value=row.descricao_soma,
                instructions=[
                    "Confirme que o campo Descrição está visível.",
                    "O valor esperado vem de CONTAORDEM[DESCRICAO SOMA].",
                    "Pressione ENTER para o Selenium preencher a Descrição.",
                ],
            )
        with step(log, "entradas_saidas.fill.descricao", row=row.row_number, tipo=row.tipo.value, field="DESCRICAO"):
            v = self._type_and_validate_candidates(
                self.DESCRICAO_CANDIDATES,
                row.descricao_soma,
                row=row,
                field="descricao",
                clear=True,
            )
            self._emit(f"Descrição preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

        if row.tipo == TipoMovimento.SAIDA:
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.FORM.VALOR",
                phase="BEFORE",
                action="INPUT",
                element_name="VALOR",
                locator=self.VALOR,
                value=row.importancia,
                instructions=[
                    "Confirme que o campo Valor está visível.",
                    "O valor esperado é a importância da linha atual.",
                    "Pressione ENTER para o Selenium preencher o Valor.",
                ],
            )
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
            if not self.a.exists(self.CAIXA_ENTRADA_CONTAINER, timeout_seconds=2):
                self._emit("Aviso: Campo de Caixa (Entrada) não encontrado no formulário.", level=logging.WARNING)
                return

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
                self._emit(f"Caixa selecionada com sucesso: {chosen}", row=row.row_number, tipo=row.tipo.value)

                if self.strict_caixa and not self._match_ok(row.caixa, chosen):
                    self._emit(
                        f"Aviso: CAIXA_ENTRADA não manteve a seleção esperada. "
                        f"esperado='{row.caixa}' | selecionado='{chosen}'",
                        level=logging.WARNING,
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
            except Exception as e:
                self._emit(
                    f"Aviso: falha ao validar CAIXA_ENTRADA, vou continuar o fluxo. erro='{e}'",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )

    def _fill_saida(self, row: ContaOrdemRow) -> None:
        self._debug_checkpoint(
            row=row,
            stage="SAIDA.FORM.DATA_VENCIMENTO",
            phase="BEFORE",
            action="INPUT",
            element_name="DATA_VENCIMENTO_SAIDA",
            locator=self.DATA_VENCIMENTO_SAIDA,
            value=row.data_mov,
            instructions=[
                "Confirme que o campo Data Vencimento está visível.",
                f"Se quiser testar o XPath, use: x {self.DATA_VENCIMENTO_SAIDA[1]}",
                "Pressione ENTER para o Selenium preencher a Data Vencimento.",
            ],
        )
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
        self._debug_checkpoint(
            row=row,
            stage="SAIDA.FORM.DATA_VENCIMENTO",
            phase="AFTER",
            action="INPUT",
            element_name="DATA_VENCIMENTO_SAIDA",
            locator=self.DATA_VENCIMENTO_SAIDA,
            value=row.data_mov,
            instructions=[
                "Confirme que a Data Vencimento ficou preenchida corretamente.",
                "O próximo passo será salvar o formulário principal.",
                "Pressione ENTER para continuar.",
            ],
        )

    def _save_form_if_present(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.save_form_best_effort", row=row.row_number, tipo=row.tipo.value):
            try:
                self._dismiss_overlays()
                self._close_datepicker()
                if self.a.exists(self.BTN_SALVAR_FORM, timeout_seconds=1):
                    try:
                        btn = self.a.driver.find_element(*self.BTN_SALVAR_FORM)
                        self.a.driver.execute_script("arguments[0].click();", btn)
                    except Exception:
                        self.a.click_js(self.BTN_SALVAR_FORM)
                    time.sleep(1)
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
            if not self.a.exists(self.BTN_REALIZAR_PAGAMENTO, timeout_seconds=1):
                self._emit(
                    "Aviso: botão 'Realizar pagamento' não ficou disponível a tempo. Vou continuar o fluxo.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return
            try:
                self.a.click_js(self.BTN_REALIZAR_PAGAMENTO)
            except TimeoutException as e:
                self._emit(
                    "Aviso: botão 'Realizar pagamento' não ficou disponível a tempo. Vou continuar o fluxo.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                log_kv(log, "Botão Realizar pagamento indisponível.", level=logging.WARNING, row=row.row_number, tipo=row.tipo.value, erro=str(e))
                return
            time.sleep(2)
            self._emit("Realizar pagamento salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)

    def _resolve_caixa_pagamento_modal_select(self) -> Any:
        return self.a.driver.find_element(*self.CAIXA_PAGAMENTO_MODAL)

    def _pagamento_saida_modal_strict(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.pagamento_saida_modal", row=row.row_number, tipo=row.tipo.value):
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.INSERIR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_INSERIR_PAGAMENTO_SAIDA",
                locator=self.BTN_INSERIR_PAGAMENTO_SAIDA,
                instructions=[
                    "Confirme que o modal da Saída principal está aberto.",
                    f"Se quiser testar o XPath, use: x {self.BTN_INSERIR_PAGAMENTO_SAIDA[1]}",
                    "Pressione ENTER para abrir o modal Inserir Pagamento.",
                ],
            )
            self._click_fixed_visible(
                self.BTN_INSERIR_PAGAMENTO_SAIDA,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.open",
                element_name="BTN_INSERIR_PAGAMENTO_SAIDA",
                timeout_seconds=20,
            )
            time.sleep(1)
            self._emit("Inserir pagamento aberto com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.INSERIR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_INSERIR_PAGAMENTO_SAIDA",
                locator=self.BTN_INSERIR_PAGAMENTO_SAIDA,
                instructions=[
                    "Confirme que o modal Inserir Pagamento foi aberto.",
                    "O próximo passo será preencher a Data Pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.DATA",
                phase="BEFORE",
                action="INPUT",
                element_name="DATA_PAGAMENTO_MODAL",
                locator=self.DATA_PAGAMENTO_MODAL,
                value=row.data_mov,
                instructions=[
                    "Confirme que o campo Data Pagamento está visível no modal.",
                    f"Se quiser testar o XPath, use: x {self.DATA_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para preencher a Data Pagamento.",
                ],
            )
            v = self._type_fixed_visible(
                self.DATA_PAGAMENTO_MODAL,
                row.data_mov,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.data_pagamento",
                element_name="DATA_PAGAMENTO_MODAL",
                timeout_seconds=20,
            )
            self._emit(f"Data de pagamento preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.DATA",
                phase="AFTER",
                action="INPUT",
                element_name="DATA_PAGAMENTO_MODAL",
                locator=self.DATA_PAGAMENTO_MODAL,
                value=row.data_mov,
                instructions=[
                    "Confirme que a Data Pagamento foi preenchida corretamente.",
                    "O próximo passo será selecionar a Forma de Pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.FORMA",
                phase="BEFORE",
                action="SELECT",
                element_name="FORMA_PAGAMENTO_MODAL",
                locator=self.FORMA_PAGAMENTO_MODAL,
                value=row.forma_pagamento,
                instructions=[
                    "Confirme que o campo Forma de Pagamento está visível.",
                    f"Se quiser testar o XPath, use: x {self.FORMA_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para selecionar a Forma de Pagamento.",
                ],
            )
            chosen_fp = self._select_fixed_visible_text(
                self.FORMA_PAGAMENTO_MODAL,
                row.forma_pagamento,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.forma_pagamento",
                element_name="FORMA_PAGAMENTO_MODAL",
                timeout_seconds=20,
            )
            self._emit(f"Forma de pagamento selecionada com sucesso: {chosen_fp}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.FORMA",
                phase="AFTER",
                action="SELECT",
                element_name="FORMA_PAGAMENTO_MODAL",
                locator=self.FORMA_PAGAMENTO_MODAL,
                value=chosen_fp,
                instructions=[
                    "Confirme que a Forma de Pagamento ficou selecionada corretamente.",
                    "O próximo passo depende do tipo de pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            if self._norm(row.forma_pagamento) == self._norm("TRANSFERÊNCIA BANCÁRIA"):
                numero_documento = self._numero_documento_para_pagamento_saida(row)
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.PAGAMENTO.NUM_DOCUMENTO",
                    phase="BEFORE",
                    action="INPUT",
                    element_name="NUM_DOCUMENTO_MODAL",
                    locator=self.NUM_DOCUMENTO_MODAL,
                    value=numero_documento,
                    instructions=[
                        "Confirme que o campo Nº Documento está visível.",
                        "O valor esperado vem de CONTAORDEM[ID_INTERNO].",
                        "Pressione ENTER para preencher o Nº Documento.",
                    ],
                )
                vdoc = self._type_fixed_visible(
                    self.NUM_DOCUMENTO_MODAL,
                    numero_documento,
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.numero_documento",
                    element_name="NUM_DOCUMENTO_MODAL",
                    timeout_seconds=20,
                )
                self._emit(f"Numero do documento preenchido com sucesso: {vdoc}", row=row.row_number, tipo=row.tipo.value)
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.PAGAMENTO.NUM_DOCUMENTO",
                    phase="AFTER",
                    action="INPUT",
                    element_name="NUM_DOCUMENTO_MODAL",
                    locator=self.NUM_DOCUMENTO_MODAL,
                    value=row.id_interno,
                    instructions=[
                        "Confirme que o Nº Documento recebeu o ID_INTERNO correto.",
                        "O próximo passo será selecionar a Caixa.",
                        "Pressione ENTER para continuar.",
                    ],
                )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CAIXA",
                phase="BEFORE",
                action="SELECT",
                element_name="CAIXA_PAGAMENTO_MODAL",
                locator=self.CAIXA_PAGAMENTO_MODAL,
                value=row.caixa,
                instructions=[
                    "Confirme que o campo Caixa está visível no modal.",
                    f"Se quiser testar o XPath, use: x {self.CAIXA_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para selecionar a Caixa.",
                ],
            )
            chosen_cx = self._select_fixed_visible_text(
                self.CAIXA_PAGAMENTO_MODAL,
                row.caixa,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.caixa",
                element_name="CAIXA_PAGAMENTO_MODAL",
                timeout_seconds=20,
                stale_retries=3,
                stale_label="CAIXA",
            )
            self._emit(f"Caixa para pagamento selecionada com sucesso: {chosen_cx}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CAIXA",
                phase="AFTER",
                action="SELECT",
                element_name="CAIXA_PAGAMENTO_MODAL",
                locator=self.CAIXA_PAGAMENTO_MODAL,
                value=chosen_cx,
                instructions=[
                    "Confirme que a Caixa ficou selecionada corretamente.",
                    "O próximo passo será salvar o pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.SALVAR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_SALVAR_PAGAMENTO_MODAL",
                locator=self.BTN_SALVAR_PAGAMENTO_MODAL,
                instructions=[
                    "Confirme que o formulário de pagamento está pronto para salvar.",
                    f"Se quiser testar o XPath, use: x {self.BTN_SALVAR_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para o Selenium clicar em Salvar Pagamento.",
                ],
            )
            self._click_fixed_visible(
                self.BTN_SALVAR_PAGAMENTO_MODAL,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.salvar",
                element_name="BTN_SALVAR_PAGAMENTO_MODAL",
                timeout_seconds=20,
            )
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.SALVAR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_SALVAR_PAGAMENTO_MODAL",
                locator=self.BTN_SALVAR_PAGAMENTO_MODAL,
                instructions=[
                    "Salvar Pagamento foi clicado.",
                    "Observe o browser e, se necessário, abra F12/DevTools.",
                    "Use x /html/body/div[5]/div/button[1], swal, buttons, events, html e url para inspecionar o popup.",
                    "Pressione ENTER para o Selenium continuar para a confirmação.",
                ],
            )
            self._emit("[SAIDA] Salvar Pagamento clicado", row=row.row_number, tipo=row.tipo.value)
            self._emit("[SAIDA] Verificando popup de confirmação do pagamento", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CONFIRMACAO",
                phase="BEFORE",
                action="WAIT/CHECK",
                element_name="OK_ALERT",
                locator=self.OK_ALERT,
                instructions=[
                    "Não clique no OK manualmente.",
                    f"Se quiser testar o XPath do popup, use: x {self.OK_ALERT[1]}",
                    "Você pode usar swal, buttons, events, html e url para inspecionar o DOM.",
                    "Pressione ENTER para o Selenium procurar o OK.",
                ],
            )

            ok_el = self._wait_fixed_displayed(
                self.OK_ALERT,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                element_name="OK_ALERT",
                timeout_seconds=35,
            )
            self._emit("[SAIDA] OK do pagamento encontrado", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CONFIRMACAO",
                phase="AFTER",
                action="WAIT/CHECK",
                element_name="OK_ALERT",
                locator=self.OK_ALERT,
                instructions=[
                    "Confirme que o popup de sucesso está visível.",
                    "Agora o Selenium vai clicar no mesmo WebElement encontrado.",
                    "Pressione ENTER para continuar.",
                ],
            )
            try:
                self.a.driver.execute_script("arguments[0].click();", ok_el)
            except Exception as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                    element_name="OK_ALERT",
                    locator=self.OK_ALERT,
                    error=exc,
                )
            self._emit("[SAIDA] OK do pagamento clicado", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CONFIRMACAO",
                phase="AFTER",
                action="CLICK",
                element_name="OK_ALERT",
                locator=self.OK_ALERT,
                instructions=[
                    "Confirme que o popup começou a fechar.",
                    "O Selenium vai aguardar o desaparecimento do popup.",
                    "Pressione ENTER para continuar.",
                ],
            )

            try:
                self.a.wait_invisible(self.OK_ALERT, timeout_seconds=10)
            except TimeoutException as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                    element_name="OK_ALERT",
                    locator=self.OK_ALERT,
                    error=exc,
                )
            self._emit("[SAIDA] Popup de confirmação fechado", row=row.row_number, tipo=row.tipo.value)

            if self.strict_caixa and not self._match_ok(row.caixa, chosen_cx):
                raise RuntimeError(f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{row.caixa}' | selecionado='{chosen_cx}'")

    def _do_baixa_saida_legacy(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.baixa", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.INSERIR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_INSERIR_PAGAMENTO_SAIDA",
                locator=self.BTN_INSERIR_PAGAMENTO_SAIDA,
                instructions=[
                    "Confirme que o modal da Saída principal está aberto.",
                    f"Se quiser testar o XPath, use: x {self.BTN_INSERIR_PAGAMENTO_SAIDA[1]}",
                    "Pressione ENTER para abrir o modal Inserir Pagamento.",
                ],
            )
            self._click_fixed_visible(
                self.BTN_INSERIR_PAGAMENTO_SAIDA,
                row=row,
                stage="entradas_saidas.baixa.open",
                element_name="BTN_INSERIR_PAGAMENTO_SAIDA",
                timeout_seconds=20,
            )
            time.sleep(1)
            self._emit("Modal de pagamento aberto com sucesso", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.INSERIR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_INSERIR_PAGAMENTO_SAIDA",
                locator=self.BTN_INSERIR_PAGAMENTO_SAIDA,
                instructions=[
                    "Confirme que o modal Inserir Pagamento foi aberto.",
                    "O próximo passo será preencher a Data Pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.DATA",
                phase="BEFORE",
                action="INPUT",
                element_name="DATA_PAGAMENTO_MODAL",
                locator=self.DATA_PAGAMENTO_MODAL,
                value=row.data_mov,
                instructions=[
                    "Confirme que o campo Data Pagamento está visível no modal.",
                    f"Se quiser testar o XPath, use: x {self.DATA_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para preencher a Data Pagamento.",
                ],
            )
            v = self._type_fixed_visible(
                self.DATA_PAGAMENTO_MODAL,
                row.data_mov,
                row=row,
                stage="entradas_saidas.baixa.data_pagamento",
                element_name="DATA_PAGAMENTO_MODAL",
                timeout_seconds=20,
            )
            self._emit(f"Data de pagamento preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.DATA",
                phase="AFTER",
                action="INPUT",
                element_name="DATA_PAGAMENTO_MODAL",
                locator=self.DATA_PAGAMENTO_MODAL,
                value=row.data_mov,
                instructions=[
                    "Confirme que a Data Pagamento foi preenchida corretamente.",
                    "O próximo passo será selecionar a Forma de Pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.FORMA",
                phase="BEFORE",
                action="SELECT",
                element_name="FORMA_PAGAMENTO_MODAL",
                locator=self.FORMA_PAGAMENTO_MODAL,
                value=row.forma_pagamento,
                instructions=[
                    "Confirme que o campo Forma de Pagamento está visível.",
                    f"Se quiser testar o XPath, use: x {self.FORMA_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para selecionar a Forma de Pagamento.",
                ],
            )
            chosen_fp = self._select_fixed_visible_text(
                self.FORMA_PAGAMENTO_MODAL,
                row.forma_pagamento,
                row=row,
                stage="entradas_saidas.baixa.forma_pagamento",
                element_name="FORMA_PAGAMENTO_MODAL",
                timeout_seconds=20,
            )
            self._emit(f"Forma de pagamento selecionada com sucesso: {chosen_fp}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.FORMA",
                phase="AFTER",
                action="SELECT",
                element_name="FORMA_PAGAMENTO_MODAL",
                locator=self.FORMA_PAGAMENTO_MODAL,
                value=chosen_fp,
                instructions=[
                    "Confirme que a Forma de Pagamento ficou selecionada corretamente.",
                    "O próximo passo depende do tipo de pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            if self._norm(row.forma_pagamento) == self._norm("TRANSFERÊNCIA BANCÁRIA"):
                numero_documento = self._numero_documento_para_pagamento_saida(row)
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.PAGAMENTO.NUM_DOCUMENTO",
                    phase="BEFORE",
                    action="INPUT",
                    element_name="NUM_DOCUMENTO_MODAL",
                    locator=self.NUM_DOCUMENTO_MODAL,
                    value=numero_documento,
                    instructions=[
                        "Confirme que o campo Nº Documento está visível.",
                        "O valor esperado vem de CONTAORDEM[ID_INTERNO].",
                        "Pressione ENTER para preencher o Nº Documento.",
                    ],
                )
                vdoc = self._type_fixed_visible(
                    self.NUM_DOCUMENTO_MODAL,
                    numero_documento,
                    row=row,
                    stage="entradas_saidas.baixa.numero_documento",
                    element_name="NUM_DOCUMENTO_MODAL",
                    timeout_seconds=20,
                )
                self._emit(f"Numero do documento preenchido com sucesso: {vdoc}", row=row.row_number, tipo=row.tipo.value)
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.PAGAMENTO.NUM_DOCUMENTO",
                    phase="AFTER",
                    action="INPUT",
                    element_name="NUM_DOCUMENTO_MODAL",
                    locator=self.NUM_DOCUMENTO_MODAL,
                    value=row.id_interno,
                    instructions=[
                        "Confirme que o Nº Documento recebeu o ID_INTERNO correto.",
                        "O próximo passo será selecionar a Caixa.",
                        "Pressione ENTER para continuar.",
                    ],
                )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CAIXA",
                phase="BEFORE",
                action="SELECT",
                element_name="CAIXA_PAGAMENTO_MODAL",
                locator=self.CAIXA_PAGAMENTO_MODAL,
                value=row.caixa,
                instructions=[
                    "Confirme que o campo Caixa está visível no modal.",
                    f"Se quiser testar o XPath, use: x {self.CAIXA_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para selecionar a Caixa.",
                ],
            )
            chosen_cx = self._select_fixed_visible_text(
                self.CAIXA_PAGAMENTO_MODAL,
                row.caixa,
                row=row,
                stage="entradas_saidas.baixa.caixa",
                element_name="CAIXA_PAGAMENTO_MODAL",
                timeout_seconds=20,
                stale_retries=3,
                stale_label="CAIXA",
            )
            self._emit(f"Caixa para pagamento selecionada com sucesso: {chosen_cx}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CAIXA",
                phase="AFTER",
                action="SELECT",
                element_name="CAIXA_PAGAMENTO_MODAL",
                locator=self.CAIXA_PAGAMENTO_MODAL,
                value=chosen_cx,
                instructions=[
                    "Confirme que a Caixa ficou selecionada corretamente.",
                    "O próximo passo será salvar o pagamento.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.SALVAR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_SALVAR_PAGAMENTO_MODAL",
                locator=self.BTN_SALVAR_PAGAMENTO_MODAL,
                instructions=[
                    "Confirme que o formulário de pagamento está pronto para salvar.",
                    f"Se quiser testar o XPath, use: x {self.BTN_SALVAR_PAGAMENTO_MODAL[1]}",
                    "Pressione ENTER para o Selenium clicar em Salvar Pagamento.",
                ],
            )
            self._click_fixed_visible(
                self.BTN_SALVAR_PAGAMENTO_MODAL,
                row=row,
                stage="entradas_saidas.baixa.salvar",
                element_name="BTN_SALVAR_PAGAMENTO_MODAL",
                timeout_seconds=20,
            )
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.SALVAR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_SALVAR_PAGAMENTO_MODAL",
                locator=self.BTN_SALVAR_PAGAMENTO_MODAL,
                instructions=[
                    "Salvar Pagamento foi clicado.",
                    "Observe o browser e, se necessário, abra F12/DevTools.",
                    "Use x /html/body/div[5]/div/button[1], swal, buttons, events, html e url para inspecionar o popup.",
                    "Pressione ENTER para o Selenium continuar para a confirmação.",
                ],
            )
            time.sleep(2)
            self._emit("[SAIDA] Salvar Pagamento clicado", row=row.row_number, tipo=row.tipo.value)
            self._emit("[SAIDA] Aguardando estabilização após Salvar Pagamento", row=row.row_number, tipo=row.tipo.value)
            time.sleep(2)
            self._emit("[SAIDA] Verificando popup de confirmação do pagamento", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CONFIRMACAO",
                phase="BEFORE",
                action="WAIT/CHECK",
                element_name="OK_ALERT",
                locator=self.OK_ALERT,
                instructions=[
                    "Não clique no OK manualmente.",
                    f"Se quiser testar o XPath do popup, use: x {self.OK_ALERT[1]}",
                    "Você pode usar swal, buttons, events, html e url para inspecionar o DOM.",
                    "Pressione ENTER para o Selenium procurar o OK.",
                ],
            )

            ok_el = self._wait_fixed_displayed(
                self.OK_ALERT,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                element_name="OK_ALERT",
                timeout_seconds=35,
            )
            self._emit("[SAIDA] OK do pagamento encontrado", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CONFIRMACAO",
                phase="AFTER",
                action="WAIT/CHECK",
                element_name="OK_ALERT",
                locator=self.OK_ALERT,
                instructions=[
                    "Confirme que o popup de sucesso está visível.",
                    "Agora o Selenium vai clicar no mesmo WebElement encontrado.",
                    "Pressione ENTER para continuar.",
                ],
            )
            try:
                self.a.driver.execute_script("arguments[0].click();", ok_el)
            except Exception as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                    element_name="OK_ALERT",
                    locator=self.OK_ALERT,
                    error=exc,
                )
            self._emit("[SAIDA] OK do pagamento clicado", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.PAGAMENTO.CONFIRMACAO",
                phase="AFTER",
                action="CLICK",
                element_name="OK_ALERT",
                locator=self.OK_ALERT,
                instructions=[
                    "Confirme que o popup começou a fechar.",
                    "O Selenium vai aguardar o desaparecimento do popup.",
                    "Pressione ENTER para continuar.",
                ],
            )

            try:
                self.a.wait_invisible(self.OK_ALERT, timeout_seconds=10)
            except TimeoutException as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                    element_name="OK_ALERT",
                    locator=self.OK_ALERT,
                    error=exc,
                )
            self._emit("[SAIDA] Popup de confirmação fechado", row=row.row_number, tipo=row.tipo.value)

            if self.strict_caixa and not self._match_ok(row.caixa, chosen_cx):
                raise RuntimeError(f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{row.caixa}' | selecionado='{chosen_cx}'")

    def _inserir_pagamento_saida(self, row: ContaOrdemRow) -> None:
        self._pagamento_saida_modal_strict(row)

    def _inserir_baixa_saida(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.baixa", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.BAIXA.INSERIR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_INSERIR_BAIXA",
                locator=self.BTN_INSERIR_BAIXA,
                instructions=[
                    "Confirme que o pagamento anterior já foi concluído.",
                    f"Se quiser testar o XPath, use: x {self.BTN_INSERIR_BAIXA[1]}",
                    "Pressione ENTER para abrir Inserir Baixa.",
                ],
            )
            self._click_fixed_visible(
                self.BTN_INSERIR_BAIXA,
                row=row,
                stage="entradas_saidas.baixa.inserir",
                element_name="BTN_INSERIR_BAIXA",
                timeout_seconds=30,
            )
            self._emit("[SAIDA][BAIXA] Inserir Baixa clicado", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.BAIXA.INSERIR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_INSERIR_BAIXA",
                locator=self.BTN_INSERIR_BAIXA,
                instructions=[
                    "Confirme que o modal Inserir Baixa foi aberto.",
                    "O próximo passo será preencher a Data da Baixa.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.BAIXA.DATA",
                phase="BEFORE",
                action="INPUT",
                element_name="DATA_BAIXA",
                locator=self.DATA_BAIXA,
                value=row.data_mov,
                instructions=[
                    "Confirme que o campo Data da Baixa está visível.",
                    f"Se quiser testar o XPath, use: x {self.DATA_BAIXA[1]}",
                    "Pressione ENTER para preencher a Data da Baixa.",
                ],
            )
            data_baixa = self._type_fixed_visible(
                self.DATA_BAIXA,
                row.data_mov,
                row=row,
                stage="entradas_saidas.baixa.data",
                element_name="DATA_BAIXA",
                timeout_seconds=30,
            )
            if not self._match_ok(row.data_mov, data_baixa):
                raise RuntimeError(
                    f"DATA_BAIXA não confirmou preenchimento. esperado='{row.data_mov}' | atual='{data_baixa}'"
                )
            self._emit(f"[SAIDA][BAIXA] Data da Baixa preenchida: {data_baixa}", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.BAIXA.DATA",
                phase="AFTER",
                action="INPUT",
                element_name="DATA_BAIXA",
                locator=self.DATA_BAIXA,
                value=data_baixa,
                instructions=[
                    "Confirme que a Data da Baixa foi preenchida corretamente.",
                    "O próximo passo será salvar a baixa.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="SAIDA.BAIXA.SALVAR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_SALVAR_BAIXA",
                locator=self.BTN_SALVAR_BAIXA,
                instructions=[
                    "Confirme que o modal Inserir Baixa está pronto para salvar.",
                    f"Se quiser testar o XPath, use: x {self.BTN_SALVAR_BAIXA[1]}",
                    "Pressione ENTER para o Selenium clicar em Salvar Baixa.",
                ],
            )
            self._click_fixed_visible(
                self.BTN_SALVAR_BAIXA,
                row=row,
                stage="entradas_saidas.baixa.salvar",
                element_name="BTN_SALVAR_BAIXA",
                timeout_seconds=30,
            )
            self._emit("[SAIDA][BAIXA] Salvar Baixa clicado", row=row.row_number, tipo=row.tipo.value)
            self._debug_checkpoint(
                row=row,
                stage="SAIDA.BAIXA.SALVAR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_SALVAR_BAIXA",
                locator=self.BTN_SALVAR_BAIXA,
                instructions=[
                    "Confirme que a baixa foi enviada para processamento.",
                    "O Selenium vai aguardar o modal desaparecer.",
                    "Pressione ENTER para continuar.",
                ],
            )
            try:
                self.a.wait_invisible(self.BTN_SALVAR_BAIXA, timeout_seconds=20)
            except TimeoutException as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.baixa.salvar",
                    element_name="BTN_SALVAR_BAIXA",
                    locator=self.BTN_SALVAR_BAIXA,
                    error=exc,
                    cause="BAIXA_NAO_CONCLUIDA",
                )
            self._emit("[SAIDA][BAIXA] Baixa concluída", row=row.row_number, tipo=row.tipo.value)

    def _do_baixa_saida_strict(self, row: ContaOrdemRow) -> None:
        self._inserir_baixa_saida(row)

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

            self.a.click_js(self.BTN_SALVAR_PAGAMENTO_MODAL)
            self._emit("[SAIDA][PAGAMENTO] Salvar Pagamento clicado", row=row.row_number, tipo=row.tipo.value)
            self._emit("[SAIDA][PAGAMENTO] Aguardando popup de confirmação do pagamento", row=row.row_number, tipo=row.tipo.value)

            ok_el = self._wait_fixed_displayed(
                self.OK_ALERT,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                element_name="OK_ALERT",
                timeout_seconds=35,
            )
            self._emit("[SAIDA][PAGAMENTO] OK do pagamento encontrado", row=row.row_number, tipo=row.tipo.value)
            try:
                self.a.driver.execute_script("arguments[0].click();", ok_el)
            except Exception as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                    element_name="OK_ALERT",
                    locator=self.OK_ALERT,
                    error=exc,
                )
            self._emit("[SAIDA][PAGAMENTO] OK do pagamento clicado", row=row.row_number, tipo=row.tipo.value)

            try:
                self.a.wait_invisible(self.OK_ALERT, timeout_seconds=10)
            except TimeoutException as exc:
                self._raise_fixed_xpath_error(
                    row=row,
                    stage="entradas_saidas.pagamento_saida_modal.confirmacao",
                    element_name="OK_ALERT",
                    locator=self.OK_ALERT,
                    error=exc,
                )
            self._emit("[SAIDA][PAGAMENTO] Popup de confirmação fechado", row=row.row_number, tipo=row.tipo.value)

            if self.strict_caixa and not self._match_ok(row.caixa, chosen_cx):
                raise RuntimeError(f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{row.caixa}' | selecionado='{chosen_cx}'")

    def _do_baixa(self, row: ContaOrdemRow) -> None:
        forma = (row.forma_pagamento or "").lower().replace(" ", "")
        if forma in {"transferência bancária".lower().replace(" ", ""), "transferenciabancaria"}:
            if not row.id_interno or not row.id_interno.strip():
                raise ValueError(f"Transferência bancária requer ID_INTERNO preenchido. Linha {row.row_number}")

        with step(log, "entradas_saidas.baixa", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._dismiss_overlays()
            log.info("[PRE-BAIXA] Tentando clicar em BTN_INSERIR_BAIXA | url=%s", self.a.driver.current_url)

            # Encontrar botão visível
            if not self._exists_any(self.BTN_INSERIR_BAIXA_CANDIDATES, timeout_seconds=2):
                log.error("[PRE-BAIXA] Nenhum candidato visível de BTN_INSERIR_BAIXA encontrado")
                self._emit(
                    "Aviso: botão de baixa não ficou visível. Vou continuar sem executar a baixa.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return
            try:
                self.a.click_any_visible(self.BTN_INSERIR_BAIXA_CANDIDATES, timeout_seconds=1)
            except TimeoutException:
                log.error("[PRE-BAIXA] Nenhum candidato visível de BTN_INSERIR_BAIXA encontrado")
                self._emit(
                    "Aviso: botão de baixa não ficou visível. Vou continuar sem executar a baixa.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            # Aguardar modal #inserirBaixa ficar visível
            time.sleep(1)
            try:
                self.a.wait_visible((By.ID, "inserirBaixa"), timeout_seconds=1)
            except TimeoutException:
                log.error("[PRE-BAIXA] Modal #inserirBaixa não ficou visível")
                self._emit(
                    "Aviso: Modal de baixa não abriu. Vou continuar sem executar a baixa.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            self._emit("Modal de baixa aberto com sucesso", row=row.row_number, tipo=row.tipo.value)

            # Preencher data dentro do modal visível
            try:
                el_date = self.a.wait_any_visible_element(
                    self.DATA_BAIXA_CANDIDATES,
                    timeout_seconds=10,
                    log_timeout=False,
                )
                el_date.clear()
                el_date.send_keys(row.data_mov)
                v = (el_date.get_attribute("value") or "").strip()
                self._emit(f"Data da baixa preenchida: {v}", row=row.row_number, tipo=row.tipo.value)
            except TimeoutException:
                log.error("[BAIXA] Campo DATA_BAIXA não encontrado ou oculto")
                raise

            # Fechar datepicker se necessário
            try:
                el_date.send_keys(Keys.TAB)
                time.sleep(0.3)
            except Exception:
                pass

            # Selecionar forma de pagamento (se necessário)
            try:
                forma_sel = self.a.wait_any_visible_element(
                    self.FORMA_PAGAMENTO_BAIXA_CANDIDATES,
                    timeout_seconds=10,
                    log_timeout=False,
                )
                forma_text = row.forma_pagamento or "TRANSFERÊNCIA BANCÁRIA"
                Select(forma_sel).select_by_visible_text(forma_text)
                self._emit(f"Forma de pagamento selecionada: {forma_text}", row=row.row_number, tipo=row.tipo.value)
            except Exception as e:
                log.warning("[BAIXA] Falha ao selecionar forma de pagamento: %s", e)

            # Preencher Nº Documento com ID_INTERNO
            try:
                el_ndoc = self.a.wait_any_visible_element(
                    self.NUM_DOCUMENTO_BAIXA_CANDIDATES,
                    timeout_seconds=5,
                    log_timeout=False,
                )
                el_ndoc.clear()
                el_ndoc.send_keys(row.id_interno or row.descricao_soma)
                self._emit(f"Nº Documento preenchido: {row.id_interno}", row=row.row_number, tipo=row.tipo.value)
            except Exception as e:
                log.warning("[BAIXA] Falha ao preencher Nº Documento: %s", e)

            # Clicar em Salvar Baixa dentro do modal visível
            try:
                self.a.click_any_visible(self.BTN_SALVAR_BAIXA_CANDIDATES, timeout_seconds=10)
                time.sleep(2)
            except Exception as e:
                log.error("[BAIXA] Falha ao clicar em Salvar Baixa: %s", e)
                raise

            self._emit("Botão Salvar Baixa clicado", row=row.row_number, tipo=row.tipo.value)

            # Aguardar modal fechar
            try:
                self.a.wait_invisible((By.ID, "inserirBaixa"), timeout_seconds=10)
            except TimeoutException:
                log.warning("[BAIXA] Modal #inserirBaixa ainda visível após Salvar")

            # Tratar SweetAlert
            self._dismiss_overlays()
            self._emit("Baixa realizada com sucesso", row=row.row_number, tipo=row.tipo.value)

    # -----------------------
    # doc search
    # -----------------------
    # -----------------------
    # doc search
    # -----------------------
    # -----------------------
    # doc search
    # -----------------------
    def _click_radio_force_change(self, locator: Tuple[str, str]) -> None:
        # a revelação do painel de data é um TOGGLE ligado ao clique do rádio,
        # não um "show" condicional: clicar de novo num retry (mesma página,
        # sem reload, painel já visível da tentativa anterior) esconde-o outra
        # vez. Por isso só clicamos se o rádio ainda não estiver marcado.
        el = self.a.wait_present(locator, timeout_seconds=5)
        if el.is_selected():
            return
        self.a.click_js(locator)

    @staticmethod
    def _is_no_results_text(text: str) -> bool:
        n = unicodedata.normalize("NFKD", (text or ""))
        n = "".join(ch for ch in n if not unicodedata.combining(ch))
        n = " ".join(n.lower().split())
        if not n:
            return False
        markers = (
            "nenhum registo encontrado",
            "nenhum registro encontrado",
            "sem resultados",
            "no matching records",
            "no data available",
            "nothing found",
        )
        return any(marker in n for marker in markers)

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

    def _search_doc_lookup_attempt(self, row: ContaOrdemRow) -> str | None:
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

        if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=2):
            return None

        try:
            doc = self.a.wait_visible(self.RESULT_DOC, timeout_seconds=5).text.strip()
        except TimeoutException as e:
            if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=2):
                return None
            raise TimeoutException(
                f"Timeout à espera do RESULT_DOC. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
            ) from e

        if not doc or self._is_no_results_text(doc):
            if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=2) or self._is_no_results_text(doc):
                return None
            raise RuntimeError("Doc ID vazio após pesquisa.")

        self._emit(f"Número do documento extraído: {doc}", row=row.row_number, tipo=row.tipo.value)
        return doc

    def search_existing_doc(self, row: ContaOrdemRow) -> str | None:
        return self._search_doc_lookup_attempt(row)

    def _search_doc_id_attempt(self, row: ContaOrdemRow) -> str:
        doc = self._search_doc_lookup_attempt(row)
        if doc is None:
            fallback_doc = f"SEM_DOC_{row.row_number}"
            log_kv(
                log,
                "Sem resultados na pesquisa do nº SOMA. Vou usar identificador provisório.",
                level=logging.WARNING,
                row=row.row_number,
                tipo=row.tipo.value,
                doc=fallback_doc,
            )
            return fallback_doc
        return doc

    def _search_doc_id(self, row: ContaOrdemRow) -> str:
        with step(log, "entradas_saidas.search_doc", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            return self._search_doc_id_attempt(row)

    def precheck_duplicate(self, row: ContaOrdemRow) -> str | None:
        with step(log, "entradas_saidas.precheck_duplicate", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            doc = self.search_existing_doc(row)
            if doc is None:
                self._emit(
                    "Nenhum registro encontrado. Pode seguir com o lançamento.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return None

            self._emit(
                f"Documento já existe. Lançamento será pulado. doc='{doc}'",
                row=row.row_number,
                tipo=row.tipo.value,
            )
            return doc

    def fetch_dados_doc(self, doc_id: str) -> str:
        url = f"{self.base_ivv}?mod=ivv&exec=entradas_saidas_dados&ID={doc_id}"
        with step(log, "entradas_saidas.fetch_dados_doc", doc=doc_id, url=url):
            self._emit(f"Redirecionando para: {url}")
            self.a.driver.get(url)
            self.a.wait_dom_ready(20)
            self._emit("Nova página carregada com sucesso!")
            try:
                cell = self.a.wait_visible(self.DADOS_DOC_CELL, timeout_seconds=5)
                txt = (cell.text or "").strip()
                self._emit(f"Número do documento extraído: {txt}")
                return txt
            except TimeoutException:
                self._emit(
                    "Aviso: não consegui ler o DADOS DOC nessa página. Vou seguir sem essa confirmação.",
                    level=logging.WARNING,
                    doc=doc_id,
                    url=url,
                )
                return ""

    def recover_doc_id(self, row: ContaOrdemRow) -> str | None:
        return self.search_existing_doc(row)

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

            # Desativa debug interativo após preencher dados
            self.a.set_debug_context(None)

            if row.tipo == TipoMovimento.SAIDA:
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.FORM.SALVAR_PRINCIPAL",
                    phase="BEFORE",
                    action="CLICK",
                    element_name="BTN_SALVAR_FORM",
                    locator=self.BTN_SALVAR_FORM,
                    instructions=[
                        "Confirme que todos os campos principais da Saída estão preenchidos.",
                        "Se quiser testar o XPath, use: x " + self.BTN_SALVAR_FORM[1],
                        "Pressione ENTER para o Selenium salvar o formulário principal.",
                    ],
                )
            self._save_form_if_present(row)
            if row.tipo == TipoMovimento.SAIDA:
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.FORM.SALVAR_PRINCIPAL",
                    phase="AFTER",
                    action="CLICK",
                    element_name="BTN_SALVAR_FORM",
                    locator=self.BTN_SALVAR_FORM,
                    instructions=[
                        "Confirme que o formulário principal foi salvo.",
                        "O próximo passo será Realizar Pagamento.",
                        "Pressione ENTER para continuar.",
                    ],
                )

            if row.tipo == TipoMovimento.SAIDA:
                if not self.a.exists(self.BTN_REALIZAR_PAGAMENTO, timeout_seconds=2):
                    raise TimeoutException("BTN_REALIZAR_PAGAMENTO não encontrado para a linha de SAÍDA.")
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.PAGAMENTO.REALIZAR",
                    phase="BEFORE",
                    action="CLICK",
                    element_name="BTN_REALIZAR_PAGAMENTO",
                    locator=self.BTN_REALIZAR_PAGAMENTO,
                    instructions=[
                        "Confirme que o botão Realizar Pagamento está visível.",
                        "O Selenium vai abrir o fluxo de pagamento da Saída.",
                        "Pressione ENTER para clicar em Realizar Pagamento.",
                    ],
                )
                self._realizar_pagamento(row)
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.PAGAMENTO.REALIZAR",
                    phase="AFTER",
                    action="CLICK",
                    element_name="BTN_REALIZAR_PAGAMENTO",
                    locator=self.BTN_REALIZAR_PAGAMENTO,
                    instructions=[
                        "Confirme que o fluxo de pagamento foi aberto com sucesso.",
                        "O próximo passo será clicar em Inserir Pagamento.",
                        "Pressione ENTER para continuar.",
                    ],
                )
                self._inserir_pagamento_saida(row)
                self._inserir_baixa_saida(row)
            else:
                if (row.forma_pagamento or "").strip().upper() == "TRANSFERÊNCIA BANCÁRIA":
                    self._realizar_pagamento(row)
                    self._do_baixa(row)

            if row.tipo == TipoMovimento.SAIDA:
                self._debug_checkpoint(
                    row=row,
                    stage="SAIDA.DOC.PESQUISA_FINAL",
                    phase="BEFORE",
                    action="SEARCH",
                    element_name="DADOS_DOC",
                    locator=self.DADOS_DOC_CELL,
                    instructions=[
                        "Confirme que a baixa/pagamento já foi concluída.",
                        "O próximo passo será pesquisar/recuperar o DOC SOMA final.",
                        "Pressione ENTER para iniciar a pesquisa final.",
                    ],
                )
                self._emit("[SAIDA][DOC] Iniciando pesquisa final", row=row.row_number, tipo=row.tipo.value)
            doc = self._search_doc_id(row)
            log_kv(log, "Documento criado.", level=logging.INFO, row=row.row_number, tipo=row.tipo.value, doc=doc)
            return doc
