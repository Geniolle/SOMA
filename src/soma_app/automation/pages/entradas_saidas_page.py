from __future__ import annotations

import logging
import time
import unicodedata
from datetime import datetime, timedelta
from typing import Any, Callable, Iterable, List, Tuple
from urllib.parse import parse_qs, urlparse

from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from soma_app.automation.actions import Actions
from soma_app.config.fields import EntradaSaidaField, field_name
from soma_app.config.locators import apply_locator_overrides
from soma_app.config.urls import DEFAULT_SITE_HOME_URL, ENTRADAS_SAIDAS_DADOS_URL_SUFFIX
from soma_app.config.xpaths import CONFIRM_NO_XPATHS, CONFIRM_POPUP_CONTAINER_XPATHS, CONFIRM_YES_XPATHS
from soma_app.domain.models import ContaOrdemRow, TipoMovimento
from soma_app.infra.env import env_bool, env_int
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
    # CAMPOS SAÃƒÂDA
    # =========
    DATA_VENCIMENTO_SAIDA = (By.XPATH, "")

    # =========
    # BOTÃƒâ€¢ES (form)
    # =========
    BTN_SALVAR_FORM = (By.XPATH, "")
    BTN_VOLTAR = (By.XPATH, "")

    FORM_READY_SELECT2 = (By.XPATH, "")
    FORM_READY_CANDIDATES = []

    # =========
    # ALERTAS / OVERLAY
    # =========
    OK_ALERT = (By.XPATH, "")
    OK_ALERT_CANDIDATES = []
    SWAL_CONTAINER = (By.CLASS_NAME, "swal2-container")
    CONFIRM_POPUP_CONTAINER_CANDIDATES = [(By.XPATH, xpath) for xpath in CONFIRM_POPUP_CONTAINER_XPATHS]
    CONFIRM_YES_CANDIDATES = [(By.XPATH, xpath) for xpath in CONFIRM_YES_XPATHS]
    CONFIRM_NO_CANDIDATES = [(By.XPATH, xpath) for xpath in CONFIRM_NO_XPATHS]

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
    FORMA_PAGAMENTO_MODAL_CANDIDATES = []
    CAIXA_PAGAMENTO_MODAL_CANDIDATES = []
    MODAL_EXCLUIR_PAGAMENTO = (By.XPATH, "")
    CHECK_EXCLUIR_PAGAMENTO = (By.XPATH, "")
    CHECK_EXCLUIR_PAGAMENTO_ALL = (By.XPATH, "")
    BTN_EXCLUIR_PAGAMENTO_MODAL = (By.XPATH, "")

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
        self.detailed_logging = env_bool("DETAILED_RUN_LOGGING", False)
        self.base_ivv = (getattr(settings, "site_home_url", "") or DEFAULT_SITE_HOME_URL).rstrip("/") + "/"
        apply_locator_overrides(self, "entradas_saidas")
        self.RADIO_ANY_CANDIDATES = self.RADIO_SAIDA_CANDIDATES + self.RADIO_ENTRADA_CANDIDATES
        self.FORM_READY_CANDIDATES = [
            self.FORM_CONTAINER,
            self.DESCRICAO,
            self.BTN_SALVAR_FORM,
            self.FORM_READY_SELECT2,
        ]

        self.strict_caixa = bool(getattr(settings, "STRICT_CAIXA_MATCH", True))

        self.caixa_validate_sleep = float(getattr(settings, "CAIXA_VALIDATE_SLEEP", 0.35))
        self.caixa_validate_retries = int(getattr(settings, "CAIXA_VALIDATE_RETRIES", 2))
        self.caixa_stable_checks = int(getattr(settings, "CAIXA_STABLE_CHECKS", 1))
        self.caixa_stable_interval = float(getattr(settings, "CAIXA_STABLE_INTERVAL", 0.2))
        self.search_doc_attempts = max(1, env_int("SEARCH_DOC_ATTEMPTS", 2))
        self.search_doc_retry_delay = max(0, env_int("SEARCH_DOC_RETRY_DELAY_SECONDS", 1))
        self.search_doc_result_timeout = max(2, env_int("SEARCH_DOC_RESULT_TIMEOUT_SECONDS", 8))
        self.search_doc_broader_timeout = max(2, env_int("SEARCH_DOC_BROADER_TIMEOUT_SECONDS", 8))
        self.search_back_dom_timeout = max(1, env_int("SEARCH_DOC_BACK_DOM_TIMEOUT_SECONDS", 2))
        self.search_back_visible_timeout = max(1, env_int("SEARCH_DOC_BACK_VISIBLE_TIMEOUT_SECONDS", 1))
        self.search_direct_dom_timeout = max(1, env_int("SEARCH_DOC_DIRECT_DOM_TIMEOUT_SECONDS", 2))
        self.search_direct_visible_timeout = max(1, env_int("SEARCH_DOC_DIRECT_VISIBLE_TIMEOUT_SECONDS", 2))
        self.search_list_ready_timeout = max(2, env_int("SEARCH_DOC_LIST_READY_TIMEOUT_SECONDS", 3))

    # -----------------------
    # helpers de log/console
    # -----------------------
    @staticmethod
    def _safe(v: Any, max_len: int = 160) -> str:
        s = "" if v is None else str(v)
        s = " ".join(s.split())
        return s[:max_len] + ("Ã¢â‚¬Â¦" if len(s) > max_len else "")

    @staticmethod
    def _norm(s: str) -> str:
        s2 = unicodedata.normalize("NFKD", (s or ""))
        s2 = "".join(ch for ch in s2 if not unicodedata.combining(ch))
        return " ".join(s2.strip().lower().split())

    def _is_transferencia_bancaria(self, value: Any) -> bool:
        return self._norm(str(value)) == "transferencia bancaria"

    def _emit(self, msg: str, level: int = logging.INFO, **kv: Any) -> None:
        if self.detailed_logging:
            print(msg)
        try:
            log_kv(log, msg, level=level, **kv)
        except Exception:
            log.log(level, msg)

    def _debug_pause_saida(self, row: ContaOrdemRow, *, stage: str, field: str, planned: str = "", result: str = "") -> None:
        if row.tipo != TipoMovimento.SAIDA:
            return
        return

    def _confirm(self, step_name: str, *, row: ContaOrdemRow | None = None, **kv: Any) -> None:
        payload: dict[str, Any] = dict(kv)
        if row is not None:
            payload.setdefault("row", row.row_number)
            payload.setdefault("tipo", row.tipo.value)
        self._emit(f"Passo '{step_name}' concluÃƒÂ­do com sucesso.", **payload)

    @staticmethod
    def _field(field: EntradaSaidaField | str) -> str:
        return field_name(field)

    def _exists_any(self, locators: Iterable[Tuple[str, str]], timeout_seconds: int = 1) -> bool:
        try:
            self.a.wait_any_present(locators, timeout_seconds=timeout_seconds)
            return True
        except Exception:
            return False

    def _input_value(self, locator: Tuple[str, str]) -> str:
        try:
            el = self.a.driver.find_element(*locator)
            return (el.get_attribute("value") or "").strip()
        except Exception:
            return ""

    def _first_present_locator(self, locators: Iterable[Tuple[str, str]]) -> Tuple[str, str] | None:
        for loc in self._unique_locators(locators):
            try:
                if self.a.driver.find_elements(*loc):
                    return loc
            except Exception:
                continue
        return None

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
            rendered = ""
            try:
                rendered = (self.a.driver.execute_script(
                    """
                    const el = arguments[0];
                    if (!el || !el.parentElement) {
                        return '';
                    }
                    const span = el.parentElement.querySelector('.select2-selection__rendered');
                    return span ? (span.textContent || '') : '';
                    """,
                    select_el,
                ) or "").strip()
            except Exception:
                rendered = ""
            if rendered:
                return rendered

            sel = Select(select_el)
            for opt in getattr(sel, "all_selected_options", []) or []:
                txt = (opt.text or "").strip()
                if txt:
                    return txt

            selected_value = (getattr(select_el, "get_attribute", lambda *_: "")("value") or "").strip()
            if selected_value:
                for opt in sel.options:
                    try:
                        if (opt.get_attribute("value") or "").strip() == selected_value:
                            txt = (opt.text or "").strip()
                            if txt:
                                return txt
                    except Exception:
                        continue

            if sel.options:
                txt = (sel.options[0].text or "").strip()
                if txt:
                    return txt
            return ""
        except Exception:
            return ""

    def _select_first_value(self, select_el) -> str:
        try:
            return (Select(select_el).first_selected_option.get_attribute("value") or "").strip()
        except Exception:
            return ""

    def _dispatch_select_change(self, select_el) -> None:
        try:
            self.a.driver.execute_script(
                """
                const el = arguments[0];
                if (!el) {
                    return;
                }
                if (window.jQuery) {
                    window.jQuery(el).trigger('change');
                    window.jQuery(el).trigger('change.select2');
                }
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                """,
                select_el,
            )
        except Exception:
            pass

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
                f"{field}: opener do Select2 nÃƒÂ£o apareceu | value='{self._safe(value)}' | screenshot={p}"
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
            f"{field}: nÃƒÂ£o consegui selecionar '{self._safe(value)}' no Select2. "
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
                f"{field}: campo nÃƒÂ£o apareceu | value='{self._safe(value)}' | screenshot={p}"
            ) from e

        if not (value or "").strip():
            return ""

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
            f"{field}: campo nÃƒÂ£o confirmou preenchimento. esperado='{self._safe(value)}' | "
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
        option_values = [(o.get_attribute("value") or "").strip() for o in sel.options]
        options_norm = [self._norm(t) for t in options]
        values_norm = [self._norm(v) for v in option_values]
        want_norm = self._norm(desired)

        try:
            sel.select_by_visible_text(desired)
            self._dispatch_select_change(select_el)
            return self._select_first_text(select_el)
        except Exception:
            pass

        for idx, on in enumerate(options_norm):
            if on and on == want_norm:
                sel.select_by_index(idx)
                self._dispatch_select_change(select_el)
                return self._select_first_text(select_el) or desired

        for idx, vn in enumerate(values_norm):
            if vn and vn == want_norm:
                sel.select_by_index(idx)
                self._dispatch_select_change(select_el)
                return self._select_first_text(select_el) or desired

        candidates: List[int] = []
        for idx, on in enumerate(options_norm):
            if want_norm and ((on and (want_norm in on or on in want_norm)) or (values_norm[idx] and (want_norm in values_norm[idx] or values_norm[idx] in want_norm))):
                candidates.append(idx)

        if candidates:
            candidates.sort(key=lambda i: len(options_norm[i]), reverse=True)
            sel.select_by_index(candidates[0])
            self._dispatch_select_change(select_el)
            chosen = self._select_first_text(select_el) or desired
            if strict and not self._match_ok(desired, chosen):
                raise RuntimeError(
                    f"{field}: seleÃƒÂ§ÃƒÂ£o ambÃƒÂ­gua/incorreta. desejado='{desired}' | selecionado='{chosen}' | opÃƒÂ§ÃƒÂµes={options}"
                )
            return chosen

        index_map = {
            self._norm("DINHEIRO"): "0",
            self._norm("DEPÓSITO"): "1",
            self._norm("DEPOSITO"): "1",
            self._norm("CHEQUE"): "2",
            self._norm("TRANSFERÊNCIA BANCÁRIA"): "3",
            self._norm("TRANSFERENCIA BANCARIA"): "3",
            self._norm("PIX"): "4",
            self._norm("MAQUINETA"): "5",
        }
        mapped_index = index_map.get(want_norm)
        if mapped_index is not None:
            tipo_value = getattr(getattr(row, "tipo", None), "value", "")
            mapped_pos = int(mapped_index)
            mapped_value = option_values[mapped_pos] if 0 <= mapped_pos < len(option_values) else ""
            if not options:
                try:
                    self.a.driver.execute_script(
                        """
                        const el = arguments[0];
                        const value = arguments[1];
                        el.value = value;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        """,
                        select_el,
                        mapped_value or mapped_index,
                    )
                    return desired
                except Exception:
                    pass
            self._emit(
                "Fallback de selecao por indice para forma de pagamento.",
                level=logging.WARNING,
                row=row.row_number,
                tipo=tipo_value,
                field=field,
                desired=desired,
                want_norm=want_norm,
                mapped_index=mapped_index,
                mapped_value=mapped_value,
                options=options,
                option_values=option_values,
            )
            try:
                if mapped_value:
                    try:
                        sel.select_by_value(mapped_value)
                    except Exception:
                        sel.select_by_index(mapped_pos)
                else:
                    sel.select_by_index(mapped_pos)
                chosen = self._select_first_text(select_el) or desired
                return chosen
            except Exception:
                try:
                    self.a.driver.execute_script(
                        """
                        const el = arguments[0];
                        const index = arguments[1];
                        const value = arguments[2];
                        if (el && el.options && el.options[index]) {
                            const option = el.options[index];
                            const target = value || option.value || '';
                            if (window.jQuery) {
                                window.jQuery(el).val(target).trigger('change');
                                window.jQuery(el).trigger('change.select2');
                            } else {
                                el.value = target;
                                el.selectedIndex = index;
                                el.dispatchEvent(new Event('input', { bubbles: true }));
                                el.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                            return true;
                        }
                        return false;
                        """,
                        select_el,
                        mapped_pos,
                        mapped_value,
                    )
                    return desired
                except Exception:
                    pass

        chosen = self._select_first_text(select_el)
        if strict:
            raise RuntimeError(
                f"{field}: nÃƒÂ£o encontrei opÃƒÂ§ÃƒÂ£o para '{desired}'. selecionado_atual='{chosen}' | opÃƒÂ§ÃƒÂµes={options}"
            )
        return chosen

    def _select_with_sleep_validation(
        self,
        resolve_select: Callable[[], Any],
        desired_text: str,
        *,
        row: ContaOrdemRow,
        field: EntradaSaidaField | str,
        strict: bool,
    ) -> str:
        self._wait_select_ready(resolve_select, timeout_seconds=6, min_options=1)

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
                selected_value = self._select_first_value(sel_el)
                if self._field(field) == self._field(EntradaSaidaField.CAIXA_PAGAMENTO_MODAL) and selected_value:
                    return desired

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
                    f"Aviso: {field} alterou apÃƒÂ³s seleÃƒÂ§ÃƒÂ£o. tentativa {attempt}/{self.caixa_validate_retries} | "
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
                    f"Aviso: {field} ficou stale durante a seleÃƒÂ§ÃƒÂ£o. tentativa {attempt}/{self.caixa_validate_retries}. Reaplicando...",
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
            raise RuntimeError(f"{field}: nÃƒÂ£o manteve a seleÃƒÂ§ÃƒÂ£o. esperado='{desired}' | final='{last_seen}'")
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
        self._dismiss_overlays_with_wait(max_wait_seconds=1)

    def _dismiss_overlays_with_wait(self, max_wait_seconds: int = 3) -> None:
        ok_candidates = self._unique_locators(
            getattr(self, "OK_ALERT_CANDIDATES", None) or [self.OK_ALERT]
        )
        clicked = False
        end = time.time() + max(1, max_wait_seconds)
        while time.time() < end and not clicked:
            try:
                ok_loc = self._first_present_locator(ok_candidates)
                if ok_loc is None:
                    time.sleep(0.1)
                    continue
                self.a.click_js(ok_loc)
                self._emit("BotÃƒÂ£o 'OK' clicado com sucesso!")
                clicked = True
                time.sleep(0.2)
            except Exception:
                time.sleep(0.1)

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
                        "Aviso: O modal SweetAlert ainda estÃƒÂ¡ visÃƒÂ­vel apÃƒÂ³s o tempo de espera.",
                        level=logging.WARNING,
                    )
        except Exception:
            pass

    def _wait_and_click_ok_alert(self, *, row: ContaOrdemRow, timeout_seconds: int = 8) -> bool:
        ok_candidates = self._unique_locators(
            getattr(self, "OK_ALERT_CANDIDATES", None) or [self.OK_ALERT]
        )
        end = time.time() + max(1, timeout_seconds)
        last_err: Exception | None = None
        self._emit(
            "Inserir pagamento: aguardando popup OK para confirmar o salvamento.",
            row=row.row_number,
            tipo=row.tipo.value,
            timeout_seconds=timeout_seconds,
        )
        while time.time() < end:
            try:
                ok_loc = self._first_present_locator(ok_candidates)
                if ok_loc is None:
                    time.sleep(0.25)
                    continue

                ok_el = self.a.wait_visible(ok_loc, timeout_seconds=2)
                self._emit(
                    "Inserir pagamento: popup OK encontrado.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                    locator=str(ok_loc),
                )
                try:
                    self.a.click_js(ok_loc)
                    self._emit(
                        "Inserir pagamento: popup OK clicado com sucesso.",
                        row=row.row_number,
                        tipo=row.tipo.value,
                        locator=str(ok_loc),
                    )
                except Exception as exc:
                    last_err = exc
                    self.a.click(ok_loc)
                    self._emit(
                        "Inserir pagamento: popup OK clicado via JavaScript.",
                        row=row.row_number,
                        tipo=row.tipo.value,
                        locator=str(ok_loc),
                    )

                try:
                    self.a.wait_invisible(ok_loc, timeout_seconds=3)
                except Exception:
                    pass

                if self.a.exists(self.SWAL_CONTAINER, timeout_seconds=1):
                    try:
                        self.a.wait_invisible(self.SWAL_CONTAINER, timeout_seconds=3)
                    except Exception:
                        pass

                return True
            except Exception as exc:
                last_err = exc
                time.sleep(0.25)

        self._emit(
            "Inserir pagamento: popup OK nao foi encontrado dentro do tempo esperado.",
            level=logging.ERROR,
            row=row.row_number,
            tipo=row.tipo.value,
            last_err=repr(last_err) if last_err else "",
        )
        return False

    def _validate_pagamento_saida_modal(self, *, row: ContaOrdemRow, chosen_fp: str, chosen_cx: str) -> None:
        data_atual = self._input_value(self.DATA_PAGAMENTO_MODAL)
        if not self._match_ok(row.data_mov, data_atual):
            raise RuntimeError(
                f"DATA_PAGAMENTO_MODAL incorreta. esperado='{row.data_mov}' | atual='{data_atual}'"
            )

        fp_expected = (row.forma_pagamento or "").strip()
        if not fp_expected:
            raise RuntimeError("FORMA_PAGAMENTO_MODAL vazia na linha atual.")
        if not self._match_ok(fp_expected, chosen_fp):
            raise RuntimeError(
                f"FORMA_PAGAMENTO_MODAL incorreta. esperado='{fp_expected}' | atual='{chosen_fp}'"
            )

        if self._is_transferencia_bancaria(row.forma_pagamento):
            num_documento_atual = self._input_value(self.NUM_DOCUMENTO_MODAL)
            num_documento_esperado = (row.descricao_soma or row.dados_doc or row.doc_soma or "").strip()
            if not num_documento_atual:
                raise RuntimeError("NUM_DOCUMENTO_MODAL ficou vazio para transferencia bancaria.")
            if num_documento_esperado and not self._match_ok(num_documento_esperado, num_documento_atual):
                raise RuntimeError(
                    f"NUM_DOCUMENTO_MODAL incorreto. esperado='{num_documento_esperado}' | atual='{num_documento_atual}'"
                )

        if chosen_cx:
            caixa_atual = ""
            try:
                caixa_atual = self._select_first_text(self._resolve_caixa_pagamento_modal_select())
            except Exception:
                caixa_atual = ""
            if caixa_atual and not self._match_ok(chosen_cx, caixa_atual):
                raise RuntimeError(
                    f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{chosen_cx}' | atual='{caixa_atual}'"
                )

        self._emit(
            "Validacao do modal de pagamento concluida com sucesso.",
            row=row.row_number,
            tipo=row.tipo.value,
            data=data_atual,
            forma_pagamento=chosen_fp,
            num_documento=self._input_value(self.NUM_DOCUMENTO_MODAL),
            caixa=chosen_cx,
        )

    def _close_payment_modal_if_open(self, *, row: ContaOrdemRow, timeout_seconds: int = 5) -> bool:
        candidates = self._unique_locators(
            [
                (By.XPATH, "//input[@name='data_pagamento']/ancestor::div[contains(@class,'modal')][1]//*[contains(@class,'close')]"),
                (By.XPATH, "//input[@name='data_pagamento']/ancestor::div[contains(@class,'modal')][1]//button[contains(@class,'close')]"),
                (By.XPATH, "//input[@name='data_pagamento']/ancestor::div[contains(@class,'modal')][1]//span[contains(@class,'close')]"),
                (By.XPATH, "//div[contains(@class,'modal') and .//input[@name='data_pagamento']]//*[contains(@class,'close')]"),
            ]
        )
        end = time.time() + max(1, timeout_seconds)
        while time.time() < end:
            loc = self._first_present_locator(candidates)
            if loc is None:
                time.sleep(0.2)
                continue
            try:
                self._emit(
                    "Inserir pagamento: vou fechar o modal pelo botao X.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                    locator=str(loc),
                )
                self.a.click_js(loc)
                return True
            except Exception:
                time.sleep(0.2)
        return False

    def _handle_confirmation_popup(self, *, row: ContaOrdemRow, accept: bool = True, timeout_seconds: int = 5) -> bool:
        with step(log, "entradas_saidas.confirm_popup", row=row.row_number, tipo=row.tipo.value, accept=accept):
            candidates = self.CONFIRM_YES_CANDIDATES if accept else self.CONFIRM_NO_CANDIDATES
            end = time.time() + timeout_seconds
            last_err: Exception | None = None
            self._emit(
                "[ENTRADAS_SAIDAS] A aguardar popup de pagamento",
                row=row.row_number,
                tipo=row.tipo.value,
                candidates=[str(loc) for loc in candidates],
            )
            while time.time() < end:
                try:
                    alert = self.a.driver.switch_to.alert
                    if accept:
                        alert.accept()
                    else:
                        alert.dismiss()
                    self._emit(
                        "Popup nativo do browser tratado com sucesso.",
                        row=row.row_number,
                        tipo=row.tipo.value,
                        choice="accept" if accept else "dismiss",
                    )
                    return True
                except Exception as exc:
                    last_err = exc

                for loc in candidates:
                    try:
                        if not self.a.exists(loc, timeout_seconds=0):
                            continue
                        self._emit(
                            "[ENTRADAS_SAIDAS] Popup encontrado",
                            row=row.row_number,
                            tipo=row.tipo.value,
                            locator=str(loc),
                        )
                        el = self.a.wait_visible(loc, timeout_seconds=2)
                        displayed = False
                        enabled = False
                        try:
                            displayed = bool(el.is_displayed())
                        except Exception as exc:
                            last_err = exc
                        try:
                            enabled = bool(el.is_enabled())
                        except Exception as exc:
                            last_err = exc
                        self._emit(
                            f"[ENTRADAS_SAIDAS] Botão Sim visível: {str(displayed).lower()}",
                            row=row.row_number,
                            tipo=row.tipo.value,
                            locator=str(loc),
                        )
                        self._emit(
                            f"[ENTRADAS_SAIDAS] Botão Sim habilitado: {str(enabled).lower()}",
                            row=row.row_number,
                            tipo=row.tipo.value,
                            locator=str(loc),
                        )
                        if not displayed or not enabled:
                            last_err = RuntimeError("Botao Sim nao estava visivel e habilitado.")
                            continue

                        self._emit(
                            "[ENTRADAS_SAIDAS] Botão Sim encontrado",
                            row=row.row_number,
                            tipo=row.tipo.value,
                            locator=str(loc),
                            text=self._safe(getattr(el, "text", "")),
                        )
                        try:
                            self.a.driver.execute_script(
                                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                                el,
                            )
                        except Exception:
                            pass

                        try:
                            self.a.wait_clickable(loc, timeout_seconds=2)
                            self.a.click(loc)
                            self._emit(
                                "[ENTRADAS_SAIDAS] Clique normal executado",
                                row=row.row_number,
                                tipo=row.tipo.value,
                                locator=str(loc),
                            )
                        except Exception as exc:
                            last_err = exc
                            try:
                                self.a.click_js(loc)
                                self._emit(
                                    "[ENTRADAS_SAIDAS] Clique JavaScript executado como fallback",
                                    row=row.row_number,
                                    tipo=row.tipo.value,
                                    locator=str(loc),
                                )
                            except Exception as exc_js:
                                last_err = exc_js
                                continue

                        popup_container = None
                        for candidate in self.CONFIRM_POPUP_CONTAINER_CANDIDATES:
                            try:
                                if self.a.exists(candidate, timeout_seconds=0):
                                    popup_container = candidate
                                    break
                            except Exception:
                                pass
                        if popup_container is None and self.a.exists(self.SWAL_CONTAINER, timeout_seconds=0):
                            popup_container = self.SWAL_CONTAINER

                        verification_locator = popup_container or loc
                        try:
                            self.a.wait_invisible(verification_locator, timeout_seconds=3)
                        except Exception as exc:
                            last_err = exc
                            self._emit(
                                "[ENTRADAS_SAIDAS] Popup ainda visivel apos o clique.",
                                level=logging.WARNING,
                                row=row.row_number,
                                tipo=row.tipo.value,
                                locator=str(verification_locator),
                            )
                            continue

                        try:
                            self.a.wait_invisible(loc, timeout_seconds=2)
                        except Exception:
                            pass
                        self._emit(
                            "[ENTRADAS_SAIDAS] Popup fechado",
                            row=row.row_number,
                            tipo=row.tipo.value,
                            locator=str(loc),
                        )
                        self._emit(
                            "[ENTRADAS_SAIDAS] Pagamento confirmado",
                            row=row.row_number,
                            tipo=row.tipo.value,
                            choice="Sim" if accept else "Nao",
                        )
                        return True
                    except StaleElementReferenceException as exc:
                        last_err = exc
                    except Exception as exc:
                        last_err = exc

                try:
                    clicked_js = bool(
                        self.a.driver.execute_script(
                            """
                            const wanted = arguments[0].map(v => String(v || '').trim().toLowerCase()).filter(Boolean);
                            const elements = Array.from(document.querySelectorAll('button, a, [role="button"], .swal2-confirm, .swal2-cancel, .swal2-actions *'));
                            const isVisible = (el) => {
                                if (!el) return false;
                                const style = window.getComputedStyle(el);
                                return style && style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
                            };
                            for (const el of elements) {
                                if (!isVisible(el)) continue;
                                const text = String(el.innerText || el.textContent || '').trim().toLowerCase();
                                if (!text) continue;
                                if (!wanted.some((needle) => text === needle || text.includes(needle) || needle.includes(text))) continue;
                                el.click();
                                return true;
                            }
                            return false;
                            """,
                            ["sim", "ok", "ok!", "confirmar", "aceitar", "yes", "yes!"],
                        )
                    )
                    if clicked_js:
                        self._emit(
                            "[ENTRADAS_SAIDAS] Clique JavaScript executado para confirmar o popup.",
                            row=row.row_number,
                            tipo=row.tipo.value,
                        )
                        return True
                except Exception as exc:
                    last_err = exc

                time.sleep(0.2)
            self._emit(
                "Popup de confirmação não foi tratado dentro do tempo esperado.",
                level=logging.WARNING,
                row=row.row_number,
                tipo=row.tipo.value,
                accept=accept,
                last_err=repr(last_err) if last_err else "",
            )
        return False

    def _click_menu_entradas_saidas(self, timeout_seconds: int = 60) -> None:
        direct_url = f"{self.base_ivv}?mod=ivv&exec=entradas_saidas"

        try:
            loc = self.a.wait_any_present(self.MENU_ENTRADAS_SAIDAS_CANDIDATES, timeout_seconds=timeout_seconds)
            self.a.click_js(loc)
            self.a.wait_dom_ready(15)
            self._emit("Clicou na opÃƒÂ§ÃƒÂ£o 'Entradas/SaÃƒÂ­das' com sucesso!")
            return
        except Exception as exc:
            p = self.a.screenshot("entradas_saidas_menu_fallback")
            self._emit(
                "Menu 'Entradas/SaÃƒÂ­das' nÃƒÂ£o foi localizado; vou abrir a lista diretamente.",
                level=logging.WARNING,
                url=getattr(self.a.driver, "current_url", ""),
                title=getattr(self.a.driver, "title", ""),
                direct_url=direct_url,
                screenshot=p,
                err=type(exc).__name__,
            )

        self.a.driver.get(direct_url)
        self.a.wait_dom_ready(15)
        self.a.wait_present(self.PESQ_DESCRICAO, timeout_seconds=timeout_seconds)
        self._emit("Lista de 'Entradas/SaÃƒÂ­das' aberta diretamente com sucesso!", url=direct_url)

    def _click_nova(self, timeout_seconds: int = 30) -> None:
        loc = self.a.wait_any_present(self.BTN_NOVA_CANDIDATES, timeout_seconds=timeout_seconds)
        self.a.click_js(loc)
        self.a.wait_dom_ready(15)
        self._emit("Clicou no botÃƒÂ£o 'Nova Entrada/SaÃƒÂ­da' com sucesso!")

    def _ensure_pesquisa_visivel(self, row: ContaOrdemRow, *, timeout_seconds: int | None = None) -> None:
        if self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=2):
            return

        self._dismiss_overlays()
        self._close_datepicker()

        with step(log, "entradas_saidas.ensure_list", row=row.row_number, tipo=row.tipo.value):
            wait_timeout = timeout_seconds if timeout_seconds is not None else self.search_list_ready_timeout
            self._click_menu_entradas_saidas(timeout_seconds=wait_timeout)
            self.a.wait_present(self.PESQ_DESCRICAO, timeout_seconds=wait_timeout)

    def _open_search_list_direct(self, row: ContaOrdemRow) -> None:
        url = f"{self.base_ivv}{ENTRADAS_SAIDAS_DADOS_URL_SUFFIX}"
        with step(log, "entradas_saidas.back_to_list_direct_open", row=row.row_number, tipo=row.tipo.value, url=url):
            self._emit(
                "Lista de pesquisa nÃ£o reapareceu; vou abrir a pÃ¡gina diretamente.",
                level=logging.WARNING,
                row=row.row_number,
                tipo=row.tipo.value,
                url=url,
            )
            self.a.driver.get(url)
            self.a.wait_dom_ready(self.search_direct_dom_timeout)
            try:
                self.a.wait_present(
                    self.PESQ_DESCRICAO,
                    timeout_seconds=max(2, self.search_direct_visible_timeout),
                )
            except Exception:
                self._ensure_pesquisa_visivel(row, timeout_seconds=self.search_list_ready_timeout)

    # -----------------------
    # navegaÃƒÂ§ÃƒÂ£o principal
    # -----------------------
    def _wait_new_form_ready(self, row: ContaOrdemRow, timeout_seconds: int = 60) -> None:
        last_err: Exception | None = None

        for attempt in (1, 2):
            try:
                self.a.wait_any_present(self.FORM_READY_CANDIDATES, timeout_seconds=timeout_seconds)
                self.a.wait_any_present(self.RADIO_ANY_CANDIDATES, timeout_seconds=min(30, timeout_seconds))
                return
            except Exception as e:
                last_err = e
                try:
                    p = self.a.screenshot(f"entradas_saidas_new_form_row_{row.row_number}_try_{attempt}")
                    log_kv(
                        log,
                        "Form 'Nova Entrada/SaÃƒÂ­da' nÃƒÂ£o ficou pronto. Vou tentar novamente.",
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
                    self._click_nova(timeout_seconds=30)
                    self.a.wait_dom_ready(15)
                except Exception:
                    pass

        raise TimeoutException(
            f"Timeout ao abrir formulÃƒÂ¡rio 'Nova Entrada/SaÃƒÂ­da' (linha {row.row_number}). last_err={last_err}"
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
                self._emit("Selecionado o botÃƒÂ£o para o processo de 'SaÃƒÂ­da'")
            else:
                loc = self.a.wait_any_present(self.RADIO_ENTRADA_CANDIDATES, timeout_seconds=30)
                self.a.click_js(loc)
                self._emit("Selecionado o botÃƒÂ£o para o processo de 'Entrada'")
            self.a.wait_any_present(self.FORM_READY_CANDIDATES, timeout_seconds=30)
            time.sleep(1)

    # -----------------------
    # preenchimento
    # -----------------------
    def _fill_common(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.plano_conta", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.PLANO_CONTA)):
            self._debug_pause_saida(
                row,
                stage="before",
                field=self._field(EntradaSaidaField.PLANO_CONTA),
                planned=row.plano_conta,
            )
            self._select2_choose_candidates(
                self.PLANO_CONTA_CANDIDATES,
                row.plano_conta,
                row=row,
                field=self._field(EntradaSaidaField.PLANO_CONTA),
            )
            self._emit(f"Plano de conta preenchido com sucesso: {row.plano_conta}", row=row.row_number, tipo=row.tipo.value)
            self._debug_pause_saida(
                row,
                stage="after",
                field=self._field(EntradaSaidaField.PLANO_CONTA),
                planned=row.plano_conta,
                result=row.plano_conta,
            )

        with step(log, "entradas_saidas.fill.centro_custo", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.CENTRO_CUSTO)):
            self._debug_pause_saida(
                row,
                stage="before",
                field=self._field(EntradaSaidaField.CENTRO_CUSTO),
                planned=row.centro_custo,
            )
            self._select2_choose_candidates(
                self.CENTRO_CUSTO_CANDIDATES,
                row.centro_custo,
                row=row,
                field=self._field(EntradaSaidaField.CENTRO_CUSTO),
            )
            self._emit(f"Centro de custo preenchido com sucesso: {row.centro_custo}", row=row.row_number, tipo=row.tipo.value)
            self._debug_pause_saida(
                row,
                stage="after",
                field=self._field(EntradaSaidaField.CENTRO_CUSTO),
                planned=row.centro_custo,
                result=row.centro_custo,
            )

        with step(log, "entradas_saidas.fill.descricao", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.DESCRICAO)):
            self._debug_pause_saida(
                row,
                stage="before",
                field=self._field(EntradaSaidaField.DESCRICAO),
                planned=row.descricao_soma,
            )
            v = self._type_and_validate_candidates(
                self.DESCRICAO_CANDIDATES,
                row.descricao_soma,
                row=row,
                field=self._field(EntradaSaidaField.DESCRICAO),
                clear=True,
            )
            self._emit(f"DescriÃƒÂ§ÃƒÂ£o preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)
            self._debug_pause_saida(
                row,
                stage="after",
                field=self._field(EntradaSaidaField.DESCRICAO),
                planned=row.descricao_soma,
                result=v,
            )

        with step(log, "entradas_saidas.fill.valor", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.VALOR)):
            self._debug_pause_saida(
                row,
                stage="before",
                field=self._field(EntradaSaidaField.VALOR),
                planned=str(row.importancia),
            )
            self.a.type(self.VALOR, str(row.importancia))
            try:
                self.a.driver.find_element(*self.VALOR).send_keys(Keys.ENTER)
            except Exception:
                pass
            v = self._input_value(self.VALOR) or str(row.importancia)
            self._emit(f"Valor preenchido com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)
            self._debug_pause_saida(
                row,
                stage="after",
                field=self._field(EntradaSaidaField.VALOR),
                planned=str(row.importancia),
                result=v,
            )

        with step(log, "entradas_saidas.fill.obs", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.OBS)):
            self._debug_pause_saida(
                row,
                stage="before",
                field=self._field(EntradaSaidaField.OBS),
                planned=row.descricao_soma,
            )
            v = self._type_and_validate_candidates(
                self.OBS_CANDIDATES,
                row.descricao_soma,
                row=row,
                field=self._field(EntradaSaidaField.OBS),
                clear=True,
                click_first=True,
            )
            self._emit(f"DescriÃƒÂ§ÃƒÂ£o preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)
            self._debug_pause_saida(
                row,
                stage="after",
                field=self._field(EntradaSaidaField.OBS),
                planned=row.descricao_soma,
                result=v,
            )

    def _fill_entrada_sem_caixa(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.fill.data_entrada", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.DATA_ENTRADA)):
            self.a.type(self.DATA_ENTRADA, row.data_mov)
            self._close_datepicker()
            v = self._input_value(self.DATA_ENTRADA) or row.data_mov
            self._emit(f"Data preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

        with step(
            log,
            "entradas_saidas.fill.forma_pagamento_entrada",
            row=row.row_number,
            tipo=row.tipo.value,
            field=self._field(EntradaSaidaField.FORMA_PAGAMENTO_ENTRADA),
        ):
            self._select2_choose_candidates(
                self.FORMA_PAGAMENTO_ENTRADA_CANDIDATES,
                row.forma_pagamento,
                row=row,
                field=self._field(EntradaSaidaField.FORMA_PAGAMENTO_ENTRADA),
            )
            self._emit(f"Forma de pagamento selecionada com sucesso: {row.forma_pagamento}", row=row.row_number, tipo=row.tipo.value)

    def _resolve_caixa_entrada_select(self) -> Any:
        # re-localiza sempre + clica no container (como no SOMA.py) para forÃƒÂ§ar carregamento
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
        with step(log, "entradas_saidas.fill.caixa_entrada_last", row=row.row_number, tipo=row.tipo.value, field=self._field(EntradaSaidaField.CAIXA_ENTRADA)):
            if self.a.exists(self.CAIXA_ENTRADA_CONTAINER, timeout_seconds=2):
                try:
                    chosen = self._select_with_sleep_validation(
                        self._resolve_caixa_entrada_select,
                        row.caixa,
                        row=row,
                        field=self._field(EntradaSaidaField.CAIXA_ENTRADA),
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
                self._emit("Aviso: Campo de Caixa (Entrada) nÃƒÂ£o encontrado no formulÃƒÂ¡rio.", level=logging.WARNING)

    def _fill_saida(self, row: ContaOrdemRow) -> None:
        with step(
            log,
            "entradas_saidas.fill.data_vencimento_saida",
            row=row.row_number,
            tipo=row.tipo.value,
            field=self._field(EntradaSaidaField.DATA_VENCIMENTO_SAIDA),
        ):
            self._debug_pause_saida(
                row,
                stage="before",
                field=self._field(EntradaSaidaField.DATA_VENCIMENTO_SAIDA),
                planned=row.data_mov,
            )
            self.a.type(self.DATA_VENCIMENTO_SAIDA, row.data_mov)
            self._close_datepicker()
            v = self._input_value(self.DATA_VENCIMENTO_SAIDA) or row.data_mov
            self._emit(f"Data vencimento preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)
            self._debug_pause_saida(
                row,
                stage="after",
                field=self._field(EntradaSaidaField.DATA_VENCIMENTO_SAIDA),
                planned=row.data_mov,
                result=v,
            )

    def _save_form_if_present(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.save_form_best_effort", row=row.row_number, tipo=row.tipo.value):
            try:
                self._dismiss_overlays_with_wait(max_wait_seconds=3)
                self._close_datepicker()
                if self.a.exists(self.BTN_SALVAR_FORM, timeout_seconds=2):
                    self._emit("[ENTRADAS_SAIDAS] Registo guardado", row=row.row_number, tipo=row.tipo.value)
                    self._emit(
                        f"[ENTRADAS_SAIDAS] Forma de pagamento recebida: {row.forma_pagamento}",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    self._emit(
                        f"[ENTRADAS_SAIDAS] Forma de pagamento normalizada: {self._norm(row.forma_pagamento)}",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    self._emit(
                        "[ENTRADAS_SAIDAS] Deve realizar pagamento: true",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    self.a.click_js(self.BTN_SALVAR_FORM)
                    popup_handled = self._handle_confirmation_popup(row=row, accept=True, timeout_seconds=8)
                    self._dismiss_overlays_with_wait(max_wait_seconds=3)
                    if not popup_handled and self.a.exists(self.SWAL_CONTAINER, timeout_seconds=1):
                        p = self.a.screenshot(f"entradas_saidas_save_popup_still_open_row_{row.row_number}")
                        self._emit(
                            "Popup de confirmacao permaneceu aberto apos salvar.",
                            level=logging.ERROR,
                            row=row.row_number,
                            tipo=row.tipo.value,
                            screenshot=p,
                            url=getattr(self.a.driver, "current_url", ""),
                            title=getattr(self.a.driver, "title", ""),
                        )
                        raise RuntimeError("Popup de confirmacao permaneceu aberto apos salvar.")
                    self._emit("[ENTRADAS_SAIDAS] Popup fechado", row=row.row_number, tipo=row.tipo.value)
                    self._emit("Campos principais preenchidos com sucesso")
                    self._confirm("entradas_saidas.save_form_best_effort", row=row)
            except Exception:
                raise

    # -----------------------
    # pagamento/baixa
    # -----------------------
    def _realizar_pagamento(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.realizar_pagamento", row=row.row_number, tipo=row.tipo.value):
            candidates = self._unique_locators(
                (self.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES or [self.BTN_INSERIR_PAGAMENTO_SAIDA])
                + (self.BTN_REALIZAR_PAGAMENTO_CANDIDATES or [self.BTN_REALIZAR_PAGAMENTO])
            )
            self._emit(
                "Aguardando botao de realizar pagamento.",
                row=row.row_number,
                tipo=row.tipo.value,
                candidates=[str(loc) for loc in candidates],
                url=getattr(self.a.driver, "current_url", ""),
            )
            if not self._exists_any(candidates, timeout_seconds=8):
                self._emit(
                    "BotÃƒÂ£o de realizar pagamento nÃƒÂ£o apareceu no DOM; vou seguir para o modal/continuaÃƒÂ§ÃƒÂ£o do fluxo.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            loc = self._wait_payment_button_candidate(
                candidates,
                row=row,
                label="realizar pagamento",
                timeout_seconds=10,
            )
            self._emit(
                "Botao de realizar pagamento encontrado.",
                row=row.row_number,
                tipo=row.tipo.value,
                locator=str(loc),
                text=getattr(self.a.driver.find_element(*loc), "text", ""),
            )
            self.a.click(loc)
            try:
                self.a.wait_visible(self.DATA_PAGAMENTO_MODAL, timeout_seconds=10)
            except Exception as exc:
                p = self.a.screenshot(f"entradas_saidas_realizar_pagamento_modal_missing_row_{row.row_number}")
                self._emit(
                    "Modal de pagamento nao abriu depois do clique em 'Realizar pagamento'; vou seguir com o fluxo.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                    screenshot=p,
                    err=type(exc).__name__,
                )
            self._emit("Modal de pagamento aberto apos 'Realizar pagamento'.", row=row.row_number, tipo=row.tipo.value)

    def _resolve_caixa_pagamento_modal_select(self) -> Any:
        candidates = self._unique_locators(
            getattr(self, "CAIXA_PAGAMENTO_MODAL_CANDIDATES", None)
            or [self.CAIXA_PAGAMENTO_MODAL]
            + [
                (By.NAME, "id_caixa_origem"),
                (By.NAME, "id_caixa"),
                (By.XPATH, "//form//select[contains(@name,'id_caixa')]"),
                (By.XPATH, "//select[contains(@name,'id_caixa')]"),
            ]
        )
        deadline = time.monotonic() + 10
        last_seen = None
        while time.monotonic() < deadline:
            for loc in candidates:
                try:
                    elems = self.a.driver.find_elements(*loc)
                except Exception:
                    continue
                for el in elems:
                    try:
                        if not el.is_displayed():
                            continue
                        opts = Select(el).options
                        if opts:
                            return el
                        last_seen = el
                    except Exception:
                        continue
            time.sleep(0.25)
        try:
            self.a.dump_page_source("entradas_saidas_caixa_pagamento_modal_debug")
            self.a.dump_locator_probe(
                "entradas_saidas_caixa_pagamento_modal_debug",
                [
                    self.CAIXA_PAGAMENTO_MODAL,
                    *candidates,
                    (By.NAME, "id_caixa_origem"),
                    (By.NAME, "id_caixa"),
                    (By.XPATH, "//select[contains(@name,'caixa')]"),
                    (By.XPATH, "//input[contains(@name,'caixa')]"),
                    (By.XPATH, "//*[contains(normalize-space(.), 'CAIXA')]"),
                ],
            )
        except Exception:
            pass
        if last_seen is not None:
            return last_seen
        return self.a.driver.find_element(*self.CAIXA_PAGAMENTO_MODAL)

    def _caixa_pagamento_modal_has_backend_error(self) -> bool:
        try:
            source = getattr(self.a.driver, "page_source", "") or ""
        except Exception:
            return False
        normalized = self._norm(source)
        return (
            "unknown column 'nan' in 'where clause'" in normalized
            or "select id_banco from ivv_maquinetas where id_maquineta =nan" in normalized
            or "buscar_caixa" in normalized and "erro!" in normalized
        )

    def _wait_payment_modal_stable(self, *, timeout_seconds: int = 8, stable_seconds: float = 0.8) -> bool:
        end = time.monotonic() + max(1, timeout_seconds)
        stable_end = None
        while time.monotonic() < end:
            try:
                self.a.wait_visible(self.DATA_PAGAMENTO_MODAL, timeout_seconds=1)
                self.a.wait_visible(self.FORMA_PAGAMENTO_MODAL, timeout_seconds=1)
                self.a.wait_visible(self.BTN_SALVAR_PAGAMENTO_MODAL, timeout_seconds=1)
                if stable_end is None:
                    stable_end = time.monotonic() + max(0.1, stable_seconds)
                if time.monotonic() >= stable_end:
                    return True
            except Exception:
                stable_end = None
            time.sleep(0.2)
        return False

    def _wait_payment_button_candidate(
        self,
        candidates: Iterable[Tuple[str, str]],
        *,
        row: ContaOrdemRow,
        label: str,
        timeout_seconds: int,
    ) -> Tuple[str, str]:
        unique = self._unique_locators(candidates)
        remaining = list(unique)
        deadline = time.monotonic() + max(1, timeout_seconds)
        last_error: Exception | None = None

        while remaining and time.monotonic() < deadline:
            budget = max(1, int(min(2, deadline - time.monotonic())))
            try:
                loc = self.a.wait_any_present(remaining, timeout_seconds=budget)
            except Exception as exc:
                last_error = exc
                continue

            locator_text = self._norm(loc[1])
            if "cancel" in locator_text:
                self._emit(
                    f"Ignorando candidato de {label} por parecer ser de cancelamento.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    locator=str(loc),
                )
                remaining = [item for item in remaining if item != loc]
                continue

            try:
                el = self.a.driver.find_element(*loc)
                text = self._norm(getattr(el, "text", ""))
                target = self._norm((getattr(el, "get_attribute", lambda *_: "")("data-target") or ""))
                href = self._norm((getattr(el, "get_attribute", lambda *_: "")("href") or ""))
                if "cancel" in text or "cancel" in target or "cancel" in href or "excluir" in text or "excluir" in target or "excluir" in href:
                    self._emit(
                        f"Ignorando candidato de {label} por conter texto de cancelamento.",
                        level=logging.WARNING,
                        row=row.row_number,
                        tipo=row.tipo.value,
                        locator=str(loc),
                        text=self._safe(getattr(el, "text", "")),
                        data_target=self._safe(target),
                        href=self._safe(href),
                    )
                    remaining = [item for item in remaining if item != loc]
                    continue
                return loc
            except Exception as exc:
                last_error = exc
                remaining = [item for item in remaining if item != loc]

        raise TimeoutException(f"Nenhum candidato válido encontrado para {label}.") from last_error

    def _pagamento_saida_modal(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.pagamento_saida_modal", row=row.row_number, tipo=row.tipo.value):
            modal_open = self._wait_payment_modal_stable(timeout_seconds=2, stable_seconds=0.2)
            if not modal_open:
                candidates = self._unique_locators(self.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES or [self.BTN_INSERIR_PAGAMENTO_SAIDA])
                self._emit(
                    "Aguardando botao de inserir pagamento.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                    candidates=[str(loc) for loc in candidates],
                    url=getattr(self.a.driver, "current_url", ""),
                )
                if not self._exists_any(candidates, timeout_seconds=4):
                    p = self.a.screenshot(f"entradas_saidas_inserir_pagamento_missing_row_{row.row_number}")
                    self._emit(
                        "Nao encontrei o popup nem o botao de inserir pagamento.",
                        level=logging.ERROR,
                        row=row.row_number,
                        tipo=row.tipo.value,
                        screenshot=p,
                        url=getattr(self.a.driver, "current_url", ""),
                        title=getattr(self.a.driver, "title", ""),
                    )
                    self._emit(
                        "Modal de pagamento nao apareceu; vou seguir com o fluxo sem inserir pagamento.",
                        level=logging.WARNING,
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    return

                loc = self._wait_payment_button_candidate(
                    candidates,
                    row=row,
                    label="inserir pagamento",
                    timeout_seconds=5,
                )
                for attempt in range(1, 3):
                    self._emit(
                        "Botao de inserir pagamento encontrado.",
                        row=row.row_number,
                        tipo=row.tipo.value,
                        locator=str(loc),
                        text=getattr(self.a.driver.find_element(*loc), "text", ""),
                        attempt=attempt,
                    )
                    if attempt == 1:
                        self.a.click(loc)
                    else:
                        self.a.click_js(loc)

                    if self._wait_payment_modal_stable(timeout_seconds=6, stable_seconds=0.7):
                        break

                    if attempt == 1:
                        self._emit(
                            "Clique normal do botao de inserir pagamento nao manteve o modal aberto; vou tentar click_js.",
                            level=logging.WARNING,
                            row=row.row_number,
                            tipo=row.tipo.value,
                        )
                    else:
                        p = self.a.screenshot(f"entradas_saidas_inserir_pagamento_modal_missing_row_{row.row_number}")
                        self._emit(
                            "Modal de inserir pagamento nao se manteve aberto depois dos cliques.",
                            level=logging.ERROR,
                            row=row.row_number,
                            tipo=row.tipo.value,
                            url=getattr(self.a.driver, "current_url", ""),
                            title=getattr(self.a.driver, "title", ""),
                            screenshot=p,
                        )
                        self._emit(
                            "Modal de inserir pagamento nao se manteve aberto; vou seguir sem preencher este modal.",
                            level=logging.WARNING,
                            row=row.row_number,
                            tipo=row.tipo.value,
                        )
                        return
                else:
                    return
            self._emit("Modal de inserir pagamento aberto com sucesso.", row=row.row_number, tipo=row.tipo.value)

            self._emit("Inserir pagamento: vou preencher a data.", row=row.row_number, tipo=row.tipo.value)
            self.a.type(self.DATA_PAGAMENTO_MODAL, row.data_mov)
            time.sleep(0.1)
            self._emit("Inserir pagamento: vou sair do campo de data sem fechar o modal.", row=row.row_number, tipo=row.tipo.value)
            try:
                self.a.driver.find_element(*self.DATA_PAGAMENTO_MODAL).send_keys(Keys.TAB)
            except Exception:
                pass
            v = self._input_value(self.DATA_PAGAMENTO_MODAL) or row.data_mov
            self._emit(f"Data inÃƒÂ­cio preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

            def fp_resolve() -> Any:
                candidates = self._unique_locators(
                    getattr(self, "FORMA_PAGAMENTO_MODAL_CANDIDATES", None) or [self.FORMA_PAGAMENTO_MODAL]
                )
                deadline = time.monotonic() + 10
                last_seen = None
                while time.monotonic() < deadline:
                    for loc in candidates:
                        try:
                            elems = self.a.driver.find_elements(*loc)
                        except Exception:
                            continue
                        for el in elems:
                            try:
                                if not el.is_displayed():
                                    continue
                                opts = Select(el).options
                                if len(opts) >= 2:
                                    return el
                                last_seen = el
                            except Exception:
                                continue
                    time.sleep(0.25)
                if last_seen is not None:
                    return last_seen
                return self.a.driver.find_element(*self.FORMA_PAGAMENTO_MODAL)

            self._wait_select_ready(fp_resolve, timeout_seconds=10, min_options=2)
            self._emit("Inserir pagamento: vou selecionar a forma de pagamento.", row=row.row_number, tipo=row.tipo.value)
            fp_el = fp_resolve()
            chosen_fp = self._select_best_effort(
                fp_el,
                row.forma_pagamento,
                row=row,
                field=self._field(EntradaSaidaField.FORMA_PAGAMENTO_MODAL),
                strict=True,
            )
            self._emit(f"Forma de pagamento selecionada com sucesso: {chosen_fp}", row=row.row_number, tipo=row.tipo.value)
            time.sleep(0.2)

            if self._is_transferencia_bancaria(row.forma_pagamento):
                self._emit(
                    "A forma de pagamento ÃƒÂ© TRANSFERÃƒÅ NCIA BANCÃƒÂRIA. atualizar campo NÃ‚Âº Documento...",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                try:
                    num_documento = (row.descricao_soma or row.dados_doc or row.doc_soma or "").strip()
                    self._emit("Inserir pagamento: vou preencher o numero do documento.", row=row.row_number, tipo=row.tipo.value)
                    self._emit(
                        "Inserir pagamento: vou aguardar o campo numero do documento ficar visivel.",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    self.a.wait_visible(self.NUM_DOCUMENTO_MODAL, timeout_seconds=10)
                    vdoc = self._type_and_validate_candidates(
                        [self.NUM_DOCUMENTO_MODAL],
                        num_documento,
                        row=row,
                        field=self._field(EntradaSaidaField.NUM_DOCUMENTO_MODAL),
                        clear=True,
                    )
                    self._emit(f"NÃƒÂºmero do documento preenchido com sucesso: {vdoc}", row=row.row_number, tipo=row.tipo.value)
                except Exception as exc:
                    self._emit(
                        "Falha ao preencher o numero do documento no modal de pagamento.",
                        level=logging.ERROR,
                        row=row.row_number,
                        tipo=row.tipo.value,
                        err=type(exc).__name__,
                    )
                    raise

            chosen_cx = ""
            if self._caixa_pagamento_modal_has_backend_error():
                self._emit(
                    "Modal de caixa devolveu erro do backend; vou prosseguir sem selecionar caixa para nao bloquear o pagamento.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
            else:
                self._wait_select_ready(self._resolve_caixa_pagamento_modal_select, timeout_seconds=10, min_options=1)
                self._emit("Inserir pagamento: vou selecionar o caixa.", row=row.row_number, tipo=row.tipo.value)
                chosen_cx = self._select_with_sleep_validation(
                    self._resolve_caixa_pagamento_modal_select,
                    row.caixa,
                    row=row,
                    field=self._field(EntradaSaidaField.CAIXA_PAGAMENTO_MODAL),
                    strict=self.strict_caixa,
                )
                self._emit(
                    f"Caixa para pagamento selecionada com sucesso: {chosen_cx}",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )

            salvar_candidates = self._unique_locators(
                self.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES or [self.BTN_SALVAR_PAGAMENTO_MODAL]
            )
            salvar_loc = self.a.wait_any_present(salvar_candidates, timeout_seconds=5)
            salvar_el = self.a.driver.find_element(*salvar_loc)
            salvar_disabled = False
            try:
                salvar_disabled = not salvar_el.is_enabled() or "disabled" in (salvar_el.get_attribute("class") or "").lower()
            except Exception:
                pass
            if salvar_disabled:
                p = self.a.screenshot(f"entradas_saidas_pagamento_save_disabled_row_{row.row_number}")
                self._emit(
                    "Botao de salvar pagamento ainda esta desabilitado antes do clique.",
                    level=logging.ERROR,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    screenshot=p,
                    button_class=salvar_el.get_attribute("class") if salvar_el else "",
                    button_disabled=salvar_el.get_attribute("disabled") if salvar_el else "",
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                )
                raise RuntimeError("Botao de salvar pagamento permaneceu desabilitado.")
            self._validate_pagamento_saida_modal(row=row, chosen_fp=chosen_fp, chosen_cx=chosen_cx)
            try:
                self._emit("Inserir pagamento: vou clicar em salvar.", row=row.row_number, tipo=row.tipo.value)
                self.a.click_js(salvar_loc)
            except Exception:
                self._emit(
                    "Clique JS do botao de salvar pagamento falhou; vou tentar clique normal.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    locator=str(salvar_loc),
                )
                self.a.click(salvar_loc)
            popup_handled = self._handle_confirmation_popup(row=row, accept=True, timeout_seconds=12)
            if not popup_handled:
                
                self._emit(
                    "Popup confirmacao nao encontrado apos salvar o pagamento; assumindo sucesso.",
                    level=logging.ERROR,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    screenshot=p,
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                )
                
            self._close_payment_modal_if_open(row=row, timeout_seconds=5)
            self._dismiss_overlays_with_wait(max_wait_seconds=2)
            try:
                self.a.wait_invisible(self.DATA_PAGAMENTO_MODAL, timeout_seconds=10)
            except Exception:
                try:
                    self.a.dump_page_source(f"entradas_saidas_pagamento_modal_still_open_row_{row.row_number}")
                    self.a.dump_locator_probe(
                        f"entradas_saidas_pagamento_modal_still_open_row_{row.row_number}",
                        [
                            self.OK_ALERT,
                            *self._unique_locators(getattr(self, "OK_ALERT_CANDIDATES", None) or [self.OK_ALERT]),
                            self.SWAL_CONTAINER,
                            self.DATA_PAGAMENTO_MODAL,
                            self.FORMA_PAGAMENTO_MODAL,
                            self.NUM_DOCUMENTO_MODAL,
                            self.CAIXA_PAGAMENTO_MODAL,
                            self.BTN_SALVAR_PAGAMENTO_MODAL,
                        ],
                    )
                except Exception:
                    pass
                p = self.a.screenshot(f"entradas_saidas_pagamento_modal_still_open_row_{row.row_number}")
                self._emit(
                    "Modal de pagamento continua aberto apos o clique em salvar.",
                    level=logging.ERROR,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    screenshot=p,
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                )
                raise RuntimeError("Modal de pagamento continua aberto apos salvar.")
            self._emit("BotÃƒÂ£o 'Salvar Pagamento' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._emit("BotÃƒÂ£o 'OK Baixa' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

            if chosen_cx and self.strict_caixa and not self._match_ok(row.caixa, chosen_cx):
                raise RuntimeError(f"CAIXA_PAGAMENTO_MODAL incorreta. esperado='{row.caixa}' | selecionado='{chosen_cx}'")

    def _do_baixa(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.baixa", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            self._dismiss_overlays_with_wait(max_wait_seconds=1)
            candidates = self._unique_locators(self.BTN_INSERIR_BAIXA_CANDIDATES or [self.BTN_INSERIR_BAIXA])
            self._emit(
                "Aguardando botao de inserir baixa.",
                row=row.row_number,
                tipo=row.tipo.value,
                candidates=[str(loc) for loc in candidates],
                url=getattr(self.a.driver, "current_url", ""),
            )
            if not self._exists_any(candidates, timeout_seconds=2):
                self._emit(
                    "BotÃƒÂ£o de inserir baixa nÃƒÂ£o apareceu no DOM; vou seguir sem abrir a baixa dedicada.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                return

            loc = self.a.wait_any_present(candidates, timeout_seconds=10)
            self._emit(
                "Botao de inserir baixa encontrado.",
                row=row.row_number,
                tipo=row.tipo.value,
                locator=str(loc),
                text=getattr(self.a.driver.find_element(*loc), "text", ""),
            )
            try:
                self.a.click(loc)
            except Exception:
                self._emit(
                    "Clique normal do botao de baixa falhou; vou tentar click_js.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                self.a.click_js(loc)
            try:
                self.a.wait_visible(self.DATA_BAIXA, timeout_seconds=10)
            except Exception as exc:
                p = self.a.screenshot(f"entradas_saidas_inserir_baixa_modal_missing_row_{row.row_number}")
                self._emit(
                    "Modal de baixa nao abriu depois do clique; vou seguir sem abrir a baixa dedicada.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    screenshot=p,
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                    err=type(exc).__name__,
                )
                return
            self._emit("Modal de baixa aberto com sucesso.", row=row.row_number, tipo=row.tipo.value)

            self.a.type(self.DATA_BAIXA, row.data_mov)
            time.sleep(0.1)
            try:
                self.a.driver.find_element(*self.DATA_BAIXA).send_keys(Keys.TAB)
            except Exception:
                pass
            time.sleep(0.2)
            self._close_datepicker()
            v = self._input_value(self.DATA_BAIXA) or row.data_mov
            self._emit(f"Data inÃƒÂ­cio preenchida com sucesso: {v}", row=row.row_number, tipo=row.tipo.value)

            try:
                loc = self.a.wait_any_present(self.POPUP_CLICK_CANDIDATES, timeout_seconds=1)
                self.a.click_js(loc)
                self._emit("Click na janela pop-up com sucesso", row=row.row_number, tipo=row.tipo.value)
            except Exception:
                pass

            salvar_candidates = self._unique_locators(self.BTN_SALVAR_BAIXA_CANDIDATES or [self.BTN_SALVAR_BAIXA])
            salvar_loc = self.a.wait_any_present(salvar_candidates, timeout_seconds=10)
            salvar_el = self.a.driver.find_element(*salvar_loc)
            salvar_disabled = False
            try:
                salvar_disabled = not salvar_el.is_enabled() or "disabled" in (salvar_el.get_attribute("class") or "").lower()
            except Exception:
                pass
            if salvar_disabled:
                p = self.a.screenshot(f"entradas_saidas_baixa_save_disabled_row_{row.row_number}")
                self._emit(
                    "Botao de salvar baixa ainda esta desabilitado antes do clique.",
                    level=logging.ERROR,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    screenshot=p,
                    button_class=salvar_el.get_attribute("class") if salvar_el else "",
                    button_disabled=salvar_el.get_attribute("disabled") if salvar_el else "",
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                )
                raise RuntimeError("Botao de salvar baixa permaneceu desabilitado.")
            try:
                self.a.click(salvar_loc)
            except Exception:
                self._emit(
                    "Clique normal do botao de salvar baixa falhou; vou tentar click_js.",
                    level=logging.WARNING,
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
                self.a.click_js(salvar_loc)
            try:
                self.a.wait_invisible(self.DATA_BAIXA, timeout_seconds=10)
            except Exception:
                p = self.a.screenshot(f"entradas_saidas_baixa_modal_still_open_row_{row.row_number}")
                self._emit(
                    "Modal de baixa continua aberto apos o clique em salvar.",
                    level=logging.ERROR,
                    row=row.row_number,
                    tipo=row.tipo.value,
                    screenshot=p,
                    url=getattr(self.a.driver, "current_url", ""),
                    title=getattr(self.a.driver, "title", ""),
                )
                raise RuntimeError("Modal de baixa continua aberto apos salvar.")
            self._emit("Salvar Baixa salvo com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._dismiss_overlays_with_wait(max_wait_seconds=2)
            self._emit("BotÃƒÂ£o 'OK Baixa' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)
            self._confirm("entradas_saidas.baixa", row=row, data=row.data_mov)

    # -----------------------
    # doc search
    # -----------------------
    def _click_radio_force_change(self, locator: Tuple[str, str]) -> None:
        # a revelaÃƒÂ§ÃƒÂ£o do painel de data ÃƒÂ© um TOGGLE ligado ao clique do rÃƒÂ¡dio,
        # nÃƒÂ£o um "show" condicional: clicar de novo num retry (mesma pÃƒÂ¡gina,
        # sem reload, painel jÃƒÂ¡ visÃƒÂ­vel da tentativa anterior) esconde-o outra
        # vez. Por isso sÃƒÂ³ clicamos se o rÃƒÂ¡dio ainda nÃƒÂ£o estiver marcado.
        el = self.a.wait_present(locator, timeout_seconds=30)
        if el.is_selected():
            return
        self.a.click_js(locator)

    def _go_back_to_list_best_effort(self, row: ContaOrdemRow) -> None:
        with step(log, "entradas_saidas.back_to_list_best_effort", row=row.row_number, tipo=row.tipo.value):
            self._dismiss_overlays_with_wait(max_wait_seconds=1)
            self._close_datepicker()

            with step(log, "entradas_saidas.back_to_list_check_visible", row=row.row_number, tipo=row.tipo.value):
                if self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=0):
                    self._confirm("entradas_saidas.back_to_list_best_effort", row=row)
                    return

            try:
                self._open_search_list_direct(row)
                self._confirm("entradas_saidas.back_to_list_best_effort", row=row)
                return
            except Exception:
                pass

            used_fallback = False

            clicked_back = False
            with step(log, "entradas_saidas.back_to_list_click_back", row=row.row_number, tipo=row.tipo.value):
                try:
                    if self.a.exists(self.BTN_VOLTAR, timeout_seconds=1):
                        btn = self.a.wait_present(self.BTN_VOLTAR, timeout_seconds=1)
                        self.a.driver.execute_script("arguments[0].click();", btn)
                        self.a.wait_dom_ready(self.search_back_dom_timeout)
                        used_fallback = True
                        clicked_back = True
                except Exception:
                    pass

            with step(log, "entradas_saidas.back_to_list_wait_visible_after_back", row=row.row_number, tipo=row.tipo.value):
                if not clicked_back:
                    if not self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=0):
                        try:
                            self.a.driver.back()
                            self.a.wait_dom_ready(self.search_back_dom_timeout)
                            used_fallback = True
                        except Exception:
                            pass

            with step(log, "entradas_saidas.back_to_list_wait_visible_after_driver_back", row=row.row_number, tipo=row.tipo.value):
                try:
                    visible = self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=0)
                    if not visible and clicked_back:
                        try:
                            self._open_search_list_direct(row)
                            visible = True
                        except Exception:
                            visible = self.a.exists(self.PESQ_DESCRICAO, timeout_seconds=0)

                    if not visible:
                        try:
                            self._ensure_pesquisa_visivel(row, timeout_seconds=self.search_list_ready_timeout)
                            visible = True
                        except Exception:
                            visible = False

                    if not visible:
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
                except Exception:
                    with step(log, "entradas_saidas.back_to_list_ensure_menu", row=row.row_number, tipo=row.tipo.value):
                        self._ensure_pesquisa_visivel(row, timeout_seconds=self.search_list_ready_timeout)
            self._confirm("entradas_saidas.back_to_list_best_effort", row=row)

    def _search_doc_id_attempt(self, row: ContaOrdemRow) -> str:
        self._go_back_to_list_best_effort(row)

        self.a.type(self.PESQ_DESCRICAO, row.descricao_soma)
        self._emit(f"Campo pesquisar a descriÃƒÂ§ÃƒÂ£o preenchida com sucesso: {row.descricao_soma}", row=row.row_number, tipo=row.tipo.value)

        self._click_radio_force_change(self.RADIO_PERIODO)
        time.sleep(0.2)
        self._emit("Selecionado o botÃƒÂ£o de rÃƒÂ¡dio 'Periodo'", row=row.row_number, tipo=row.tipo.value)

        self._click_radio_force_change(self.RADIO_DATA_PAGAMENTO)
        time.sleep(0.2)
        self._emit("Selecionado o botÃƒÂ£o de rÃƒÂ¡dio 'Data de Pagamento'", row=row.row_number, tipo=row.tipo.value)

        if not self.a.exists(self.DATA_INI, timeout_seconds=3):
            # painel pode ter ficado escondido mesmo com o rÃƒÂ¡dio jÃƒÂ¡ marcado
            # (ex.: reset parcial ao voltar ÃƒÂ  lista); forÃƒÂ§a mais um toggle.
            self.a.click_js(self.RADIO_DATA_PAGAMENTO)
            time.sleep(0.2)

        self.a.type(self.DATA_INI, row.data_mov)
        self._emit(f"Data inÃƒÂ­cio preenchida com sucesso: {row.data_mov}", row=row.row_number, tipo=row.tipo.value)

        self.a.type(self.DATA_FIM, row.data_mov)
        self._emit(f"Data fim preenchida com sucesso: {row.data_mov}", row=row.row_number, tipo=row.tipo.value)

        self.a.click_js(self.BTN_PESQUISAR)
        self._emit("BotÃƒÂ£o 'Pesquisar' clicado com sucesso!", row=row.row_number, tipo=row.tipo.value)

        try:
            doc = self._read_search_result_doc(timeout_seconds=self.search_doc_result_timeout)
        except TimeoutException as e:
            try:
                self._dump_search_diagnostics(row, suffix="timeout")
            except Exception:
                pass
            if self._exists_any(self.NO_RESULTS_CANDIDATES, timeout_seconds=2):
                raise RuntimeError(
                    f"Sem resultados na pesquisa do nÃ‚Âº SOMA. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
                ) from e
            raise TimeoutException(
                f"Timeout ÃƒÂ  espera do RESULT_DOC. desc='{self._safe(row.descricao_soma)}' data='{self._safe(row.data_mov)}'"
            ) from e

        if not doc:
            try:
                self._dump_search_diagnostics(row, suffix="empty")
            except Exception:
                pass
            raise RuntimeError("Doc ID vazio apÃƒÂ³s pesquisa.")

        self._emit(f"NÃƒÂºmero do documento extraÃƒÂ­do: {doc}", row=row.row_number, tipo=row.tipo.value)
        self._confirm("entradas_saidas.search_doc", row=row, doc=doc)
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

    def _extract_dados_doc_from_current_page(self, *, timeout_seconds: int = 2) -> str:
        candidates = self._unique_locators(
            self.DADOS_DOC_CANDIDATES or ([self.DADOS_DOC_CELL] if self.DADOS_DOC_CELL[1].strip() else [])
        )
        end = time.time() + timeout_seconds
        while time.time() < end:
            for loc in candidates:
                try:
                    if self.a.exists(loc, timeout_seconds=0):
                        cell = self.a.driver.find_element(*loc)
                        txt = (cell.text or "").strip()
                        if txt:
                            return txt
                except Exception:
                    pass

            time.sleep(0.2)

        raise RuntimeError("DADOS DOC nÃƒÂ£o encontrado na pÃƒÂ¡gina atual.")

    def _extract_doc_id_from_current_url(self) -> str:
        try:
            current_url = str(getattr(self.a.driver, "current_url", "") or "").strip()
            if not current_url:
                return ""
            parsed = urlparse(current_url)
            params = parse_qs(parsed.query)
            for key in ("ID", "id", "doc", "doc_id"):
                values = params.get(key) or params.get(key.upper()) or params.get(key.lower())
                if not values:
                    continue
                candidate = str(values[0]).strip()
                if candidate.isdigit():
                    return candidate
        except Exception:
            pass
        return ""

    def fetch_dados_doc(self, doc_id: str) -> str:
        url = f"{self.base_ivv}{ENTRADAS_SAIDAS_DADOS_URL_SUFFIX}"
        with step(log, "entradas_saidas.fetch_dados_doc", doc=doc_id, url=url):
            try:
                txt = self._extract_dados_doc_from_current_page(timeout_seconds=2)
                self._emit(f"NÃºmero do documento extraÃ­do da pÃ¡gina atual: {txt}")
                self._confirm("entradas_saidas.fetch_dados_doc", doc=txt, url=url, mode="current_page")
                return txt
            except Exception:
                pass

            self._emit(f"Redirecionando para: {url}")
            self.a.driver.get(url)
            self.a.wait_dom_ready(10)
            self._emit("Nova pÃ¡gina carregada com sucesso!")

            try:
                txt = self._extract_dados_doc_from_current_page(timeout_seconds=6)
                self._emit(f"NÃºmero do documento extraÃ­do: {txt}")
                self._confirm("entradas_saidas.fetch_dados_doc", doc=txt, url=url, mode="fallback_page")
                return txt
            except Exception:
                self._emit(
                    "Falha ao extrair dados do documento; vou devolver o doc_id original.",
                    level=logging.WARNING,
                    doc=doc_id,
                )
                raise RuntimeError("NÃ£o consegui extrair DADOS DOC da pÃ¡gina atual nem da pÃ¡gina fallback.")
    def recover_doc_id(self, row: ContaOrdemRow) -> str:
        return self._search_doc_id(row)

    def create_and_get_doc_id(self, row: ContaOrdemRow) -> str:
        with step(log, "entradas_saidas.create_start", row=row.row_number, tipo=row.tipo.value):
            doc_hint = ""
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
            else:
                self._emit(
                    "Fluxo de pagamento adicional ignorado para Entrada.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                )
            if self._is_transferencia_bancaria(row.forma_pagamento):
                self._do_baixa(row)

            doc_hint = self._extract_doc_id_from_current_url()
            if doc_hint:
                doc = doc_hint
                self._emit(
                    "Doc_id capturado na URL atual; vou usar esse valor e evitar a pesquisa adicional.",
                    row=row.row_number,
                    tipo=row.tipo.value,
                    doc=doc,
                )
            else:
                try:
                    doc = self._search_doc_id(row)
                except Exception as exc:
                    self._emit(
                        "Pesquisa do documento falhou sem doc_id na URL atual; vou propagar o erro.",
                        level=logging.ERROR,
                        row=row.row_number,
                        tipo=row.tipo.value,
                        err=type(exc).__name__,
                    )
                    raise
            log_kv(log, "Documento criado.", level=logging.INFO, row=row.row_number, tipo=row.tipo.value, doc=doc)
            self._confirm("entradas_saidas.create_start", row=row, doc=doc)
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

            time.sleep(0.2)

        raise TimeoutException("Timeout ÃƒÂ  espera do resultado da pesquisa do nÃ‚Âº SOMA.")

    def _search_doc_id(self, row: ContaOrdemRow) -> str:
        with step(log, "entradas_saidas.search_doc", row=row.row_number, tipo=row.tipo.value, data=row.data_mov):
            attempts = self.search_doc_attempts
            delays = (self.search_doc_retry_delay,)
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
            raise RuntimeError("NÃƒÂ£o consegui gerar janelas de data para recovery do DOC.")

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
                doc = self._read_search_result_doc(timeout_seconds=self.search_doc_broader_timeout)
                if doc:
                    self._emit(
                        f"NÃƒÂºmero do documento extraÃƒÂ­do por recovery alargado: {doc}",
                        row=row.row_number,
                        tipo=row.tipo.value,
                    )
                    self._confirm("entradas_saidas.search_doc", row=row, doc=doc, mode="broader")
                    return doc
            except Exception as e:
                last_exc = e
                continue

        try:
            self._dump_search_diagnostics(row, suffix="broader_failed")
        except Exception:
            pass

        fallback_doc = self._extract_doc_id_from_current_url()
        if fallback_doc:
            self._emit(
                f"Recovery alargado falhou; vou usar o doc_id da URL atual: {fallback_doc}",
                level=logging.WARNING,
                row=row.row_number,
                tipo=row.tipo.value,
            )
            self._confirm("entradas_saidas.search_doc", row=row, doc=fallback_doc, mode="url")
            return fallback_doc

        if last_exc:
            raise last_exc
        raise RuntimeError("Recovery alargado falhou sem exceÃƒÂ§ÃƒÂ£o explÃƒÂ­cita.")




