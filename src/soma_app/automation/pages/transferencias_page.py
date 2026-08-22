from __future__ import annotations

import logging
import re
import time
import unicodedata
from pathlib import Path
from typing import Any, List, Tuple

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from soma_app.automation.actions import Actions
from soma_app.automation.debug_session import GuidedDebugSession
from soma_app.config.locators import apply_locator_overrides
from soma_app.domain.models import ContaOrdemRow, format_amount_for_input
from soma_app.infra.trace import log_kv, step

log = logging.getLogger("soma_app.pages.transferencias")

Locator = Tuple[str, str]


class TransferenciaDuplicadaError(RuntimeError):
    def __init__(self, message: str, *, matching_row: str = "") -> None:
        super().__init__(message)
        self.matching_row = matching_row


class TransferenciasPage:
    """
    Fluxo de TransferÃªncia conforme SOMA.py:
      - Caixas/Bancos
      - Nova TransferÃªncia
      - Caixa SaÃ­da, Valor, Caixa Entrada, Data, DescriÃ§Ã£o
      - Salvar, OK, Voltar
    """

    MENU_CAIXAS_BANCOS_CANDIDATES: List[Locator] = []
    BTN_NOVA_TRANSFERENCIA_CANDIDATES: List[Locator] = []
    CAIXA_SAIDA_CANDIDATES: List[Locator] = []
    CENTRO_CUSTO_SAIDA_CANDIDATES: List[Locator] = []
    CAIXA_ENTRADA_CANDIDATES: List[Locator] = []
    CENTRO_CUSTO_ENTRADA_CANDIDATES: List[Locator] = []
    VALOR_ENTRADA_CANDIDATES: List[Locator] = []
    BTN_SALVAR_CANDIDATES: List[Locator] = []

    CAIXA_SAIDA = (By.XPATH, "")
    CENTRO_CUSTO_SAIDA = (By.XPATH, "")
    DATA_INICIO = (By.XPATH, "")
    DATA_FIM = (By.XPATH, "")
    BTN_PESQUISAR = (By.XPATH, "")
    VALOR = (By.XPATH, "")
    CAIXA_ENTRADA = (By.XPATH, "")
    CENTRO_CUSTO_ENTRADA = (By.XPATH, "")
    VALOR_ENTRADA = (By.XPATH, "")
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
        self.debug_session = GuidedDebugSession(actions, settings)
        self.home_url = (getattr(settings, "site_home_url", "") or "https://verbodavida.info/IVV/").strip()
        self.timeout = int(getattr(settings, "timeout_seconds", 20) or 20)
        apply_locator_overrides(self, "transferencias")

    # ###################################################################################
    # (1) DEBUG GUIADO E CONTEXTO DE SELETORES
    # ###################################################################################
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
        instructions: List[str] | None = None,
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

    def _set_debug_context(self, context: str | None) -> None:
        try:
            self.a.set_debug_context(context)
        except Exception:
            pass

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

    def _close_datepicker(self) -> None:
        try:
            self.a.driver.find_element(By.TAG_NAME, "body").click()
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

    @staticmethod
    def _transfer_match_text(value: str) -> str:
        txt = unicodedata.normalize("NFKD", value or "")
        txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
        txt = txt.lower()
        txt = re.sub(r"\[[^\]]*\]", " ", txt)
        txt = re.sub(r"\([^\)]*\)", " ", txt)
        txt = re.sub(r"[^a-z0-9,./-]+", " ", txt)
        txt = txt.replace("-", " ")
        txt = " ".join(txt.split())
        return txt

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

    def _read_any_select2_rendered_texts_v1(self) -> List[str]:
        texts: List[str] = []
        try:
            nodes = self.a.driver.find_elements(By.CSS_SELECTOR, "span.select2-selection__rendered")
        except Exception:
            nodes = []

        for node in nodes:
            txt = (node.text or "").strip()
            if txt:
                texts.append(txt)

        return texts

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
                f"{field} nÃ£o encontrou opÃ§Ã£o compatÃ­vel com '{v}' (linha {row.row_number}). "
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
                f"{field} nÃ£o encontrou opÃ§Ã£o compatÃ­vel com '{v}' (linha {row.row_number}). "
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
                    f"{field} nÃ£o foi selecionado (linha {row.row_number}). "
                    f"Esperado conter '{v}', mas ficou '{txt}'."
                )
        except Exception:
            p = self.a.screenshot(f"transfer_select2_fail_{field.lower().replace(' ', '_')}_row_{row.row_number}")
            log_kv(
                log,
                "Select2 nÃ£o confirmou seleÃ§Ã£o.",
                level=logging.ERROR,
                field=field,
                row=row.row_number,
                value=v,
                url=self.a.driver.current_url,
                screenshot=p,
            )
            raise


    def _select2_choose_verified_v2(self, opener: Locator, value: str, *, row: ContaOrdemRow, field: str) -> None:
        v = (value or '').strip()
        if not v:
            raise ValueError(f"{field} vazio na sheet (linha {row.row_number}). Preenche a coluna correta.")

        self._dismiss_alerts()

        try:
            self.a.select2_choose(opener, v)
        except Exception:
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
                    f"{field} nÃ£o encontrou opÃ§Ã£o compatÃ­vel com '{v}' (linha {row.row_number}). "
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
                if (opt.text or '').strip() == best_text:
                    best = opt
                    break

            if best is None:
                raise RuntimeError(
                    f"{field} nÃ£o encontrou opÃ§Ã£o compatÃ­vel com '{v}' (linha {row.row_number}). "
                    f"Amostra={options[:10]}"
                )

            try:
                self.a.driver.execute_script('arguments[0].click();', best)
            except Exception:
                best.click()

        time.sleep(0.8)

        rendered_texts = self._read_any_select2_rendered_texts_v1()
        want_norm = self._norm_text(v)
        if not any(want_norm in self._norm_text(txt) or self._norm_text(txt) in want_norm for txt in rendered_texts):
            time.sleep(1.2)
            rendered_texts = self._read_any_select2_rendered_texts_v1()

        if not any(want_norm in self._norm_text(txt) or self._norm_text(txt) in want_norm for txt in rendered_texts):
            p = self.a.screenshot(f"transfer_select2_fail_{field.lower().replace(' ', '_')}_row_{row.row_number}")
            log_kv(
                log,
                "Select2 nÃ£o confirmou seleÃ§Ã£o.",
                level=logging.ERROR,
                field=field,
                row=row.row_number,
                value=v,
                rendered_texts=rendered_texts,
                url=self.a.driver.current_url,
                screenshot=p,
            )
            raise RuntimeError(
                f"{field} nÃ£o foi selecionado (linha {row.row_number}). "
                f"Esperado conter '{v}', mas ficou '{rendered_texts[:5]}'."
            )

    def _select2_choose_verified_candidates(
        self,
        openers: List[Locator],
        value: str,
        *,
        row: ContaOrdemRow,
        field: str,
    ) -> None:
        last_error: Exception | None = None
        tried: List[str] = []
        fallback_openers: List[Locator] = [
            (By.XPATH, f"(//form//span[contains(@class,'select2-selection')])[{idx}]")
            for idx in range(1, 9)
        ]
        fallback_openers.extend(
            [
                (By.XPATH, f"(//form//span[contains(@class,'select2-container')])[{idx}]")
                for idx in range(1, 9)
            ]
        )

        for opener in list(openers) + fallback_openers:
            tried.append(f"{opener[0]}:{opener[1]}")
            try:
                self._select2_choose_verified_v2(opener, value, row=row, field=field)
                return
            except Exception as exc:
                last_error = exc
                log_kv(
                    log,
                    "Select2 candidato rejeitado.",
                    level=logging.WARNING,
                    field=field,
                    row=row.row_number,
                    value=value,
                    opener=opener,
                    error=str(exc),
                )

        tried_text = ", ".join(tried)
        if last_error is not None:
            raise RuntimeError(
                f"{field} nÃ£o encontrou opÃ§Ã£o compatÃ­vel com '{value}' (linha {row.row_number}). "
                f"Seletores testados: {tried_text}"
            ) from last_error
        raise RuntimeError(
            f"{field} nÃ£o encontrou opÃ§Ã£o compatÃ­vel com '{value}' (linha {row.row_number}). "
            f"Seletores testados: {tried_text}"
        )

    @staticmethod
    def _norm_amount_v1(value: Any) -> str:
        text = "" if value is None else str(value).strip()
        text = text.replace(" ", "")
        if not text:
            return ""
        if "," in text and "." in text:
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        else:
            text = text.replace(",", ".")
        try:
            amount = float(text)
        except ValueError:
            return TransferenciasPage._norm_text(text)
        return f"{amount:.2f}".replace(".", ",")

    def _collect_visible_table_rows_v1(self) -> List[Any]:
        try:
            return self.a.driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        except Exception:
            return []

    def _row_text_v1(self, element: Any) -> str:
        try:
            return " ".join((element.text or "").split())
        except Exception:
            return ""

    def _dump_transfer_audit_snapshot_v1(self, row: ContaOrdemRow, stage: str, rows_text: List[str], page_text: str = "") -> None:
        try:
            html_path = self.a.dump_page_source(f"transfer_{stage.lower()}_row_{row.row_number}")
        except Exception:
            html_path = None

        try:
            diag_dir = Path("C:/workspace/SOMA/artifacts/diagnostics")
            diag_dir.mkdir(parents=True, exist_ok=True)
            text_path = diag_dir / f"transfer_{stage.lower()}_row_{row.row_number}.txt"
            lines = [
                f"stage={stage}",
                f"row={row.row_number}",
                f"caixa_saida={row.caixa_saida}",
                f"caixa_entrada={row.caixa}",
                f"data={row.data_mov}",
                f"valor={format_amount_for_input(row.importancia)}",
                f"html={html_path or ''}",
                "rows:",
            ]
            lines.extend(f"- {item}" for item in rows_text)
            if page_text:
                lines.append("")
                lines.append("page_text:")
                lines.append(page_text[:8000])
            text_path.write_text("\n".join(lines), encoding="utf-8")
        except Exception:
            pass

    def _row_matches_transfer_v1(self, row_text: str, transfer_row: ContaOrdemRow) -> bool:
        text = self._norm_text(row_text)
        if not text:
            return False

        expected = [
            self._transfer_match_text(transfer_row.caixa_saida),
            self._transfer_match_text(transfer_row.caixa),
            self._norm_text(transfer_row.data_mov),
            self._norm_amount_v1(transfer_row.importancia),
        ]

        return all(part and part in text for part in expected)

    def _row_matches_transfer_history_v1(self, row_text: str, transfer_row: ContaOrdemRow) -> bool:
        text = self._norm_text(row_text)
        if not text:
            return False

        caixa_saida = self._transfer_match_text(transfer_row.caixa_saida)
        caixa_entrada = self._transfer_match_text(transfer_row.caixa)
        data_mov = self._norm_text(transfer_row.data_mov)
        amount = self._norm_amount_v1(transfer_row.importancia)

        if "registrado" not in text:
            return False

        if "transferencia bancaria" not in text and "transferencia" not in text:
            return False

        if data_mov and data_mov not in text:
            return False

        caixa_entrada_ok = bool(caixa_entrada and caixa_entrada in text)
        caixa_saida_ok = bool(caixa_saida and caixa_saida in text)
        amount_ok = bool(amount and amount in text)

        if caixa_saida_ok and caixa_entrada_ok and amount_ok:
            return True

        if caixa_saida_ok and caixa_entrada_ok and ("valor" in text or "pagamento" in text):
            return True

        if caixa_saida_ok and caixa_entrada_ok:
            return True

        return False

    def _audit_existing_transfer_before_new_v1(self, row: ContaOrdemRow) -> str | None:
        self._debug_checkpoint(
            row=row,
            stage="TRANSFERENCIA.AUDITORIA.LISTAGEM",
            phase="BEFORE",
            action="INPUT",
            element_name="DATA_INICIO",
            locator=self.DATA_INICIO,
            value=row.data_mov,
            instructions=[
                "Confirme que a listagem de transferÃªncias estÃ¡ visÃ­vel.",
                "Pressione ENTER para preencher a data inicial.",
            ],
        )

        self._set_debug_context("input_dados")
        self.a.type(self.DATA_INICIO, row.data_mov)
        self._debug_checkpoint(
            row=row,
            stage="TRANSFERENCIA.AUDITORIA.LISTAGEM",
            phase="AFTER",
            action="INPUT",
            element_name="DATA_INICIO",
            locator=self.DATA_INICIO,
            value=row.data_mov,
            instructions=[
                "Confirme que a data inicial ficou preenchida.",
                "O prÃ³ximo passo Ã© preencher a data final.",
                "Pressione ENTER para continuar.",
            ],
        )

        self._debug_checkpoint(
            row=row,
            stage="TRANSFERENCIA.AUDITORIA.LISTAGEM",
            phase="BEFORE",
            action="INPUT",
            element_name="DATA_FIM",
            locator=self.DATA_FIM,
            value=row.data_mov,
            instructions=[
                "Confirme que o campo de data final estÃ¡ visÃ­vel.",
                "Pressione ENTER para preencher a data final.",
            ],
        )
        self.a.type(self.DATA_FIM, row.data_mov)
        self._debug_checkpoint(
            row=row,
            stage="TRANSFERENCIA.AUDITORIA.LISTAGEM",
            phase="AFTER",
            action="INPUT",
            element_name="DATA_FIM",
            locator=self.DATA_FIM,
            value=row.data_mov,
            instructions=[
                "Confirme que a data final ficou preenchida.",
                "O prÃ³ximo passo Ã© pesquisar a listagem.",
                "Pressione ENTER para continuar.",
            ],
        )

        self._debug_checkpoint(
            row=row,
            stage="TRANSFERENCIA.AUDITORIA.LISTAGEM",
            phase="BEFORE",
            action="CLICK",
            element_name="BTN_PESQUISAR",
            locator=self.BTN_PESQUISAR,
            instructions=[
                "Confirme que os filtros de data estÃ£o prontos.",
                "Pressione ENTER para clicar em Pesquisar.",
            ],
        )
        self.a.click_js(self.BTN_PESQUISAR)
        self.a.wait_dom_ready(15)
        time.sleep(1)

        self._debug_checkpoint(
            row=row,
            stage="TRANSFERENCIA.AUDITORIA.LISTAGEM",
            phase="AFTER",
            action="CLICK",
            element_name="BTN_PESQUISAR",
            locator=self.BTN_PESQUISAR,
            value=row.data_mov,
            instructions=[
                "Confirme que a pesquisa terminou.",
                "Agora o fluxo vai comparar as linhas retornadas com a transferÃªncia atual.",
                "Pressione ENTER para continuar.",
            ],
        )

        self._close_datepicker()
        time.sleep(1)

        for element in self._collect_visible_table_rows_v1():
            row_text = self._row_text_v1(element)
            if not row_text:
                continue
            if self._row_matches_transfer_v1(row_text, row):
                return row_text

        try:
            page_text = self._norm_text(self.a.driver.page_source)
        except Exception:
            page_text = ""

        self._dump_transfer_audit_snapshot_v1(row, "list", [self._row_text_v1(el) for el in self._collect_visible_table_rows_v1() if self._row_text_v1(el)], page_text)

        if page_text:
            if self._row_matches_transfer_v1(page_text, row):
                return " | ".join(
                    part
                    for part in (
                        self._norm_text(row.caixa_saida),
                        self._norm_text(row.caixa),
                        self._norm_text(row.data_mov),
                        self._norm_amount_v1(row.importancia),
                    )
                    if part
                )

        return None

    def _audit_existing_transfer_in_form_v1(self, row: ContaOrdemRow) -> str | None:
        try:
            self.a.wait_dom_ready(15)
            time.sleep(1)
        except Exception:
            pass

        self._close_datepicker()
        time.sleep(1)

        amount = self._norm_amount_v1(row.importancia)
        caixa_saida = self._norm_text(row.caixa_saida)
        caixa_entrada = self._norm_text(row.caixa)
        data_mov = self._norm_text(row.data_mov)
        rows_text: List[str] = []

        for element in self._collect_visible_table_rows_v1():
            row_text = self._row_text_v1(element)
            if not row_text:
                continue
            rows_text.append(row_text)
            if self._row_matches_transfer_history_v1(row_text, row):
                return row_text

        try:
            page_text = self._norm_text(self.a.driver.page_source)
        except Exception:
            page_text = ""

        self._dump_transfer_audit_snapshot_v1(row, "form", rows_text, page_text)

        if page_text:
            if "registrado" in page_text and "transferencia bancaria" in page_text:
                if data_mov and data_mov in page_text and caixa_saida and caixa_entrada and caixa_saida in page_text and caixa_entrada in page_text:
                    if amount and amount in page_text:
                        return " | ".join(part for part in (caixa_saida, caixa_entrada, data_mov, amount) if part)
                    return " | ".join(part for part in (caixa_saida, caixa_entrada, data_mov) if part)

        return None

    def _open_caixas_bancos_menu_v1(self, row: ContaOrdemRow) -> None:
        self._goto_home()

        with step(log, "transfer.open_menu", row=row.row_number):
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.MENU.CAIXAS_BANCOS",
                phase="BEFORE",
                action="CLICK",
                element_name="MENU_CAIXAS_BANCOS",
                locator=self.MENU_CAIXAS_BANCOS_CANDIDATES[0] if self.MENU_CAIXAS_BANCOS_CANDIDATES else None,
                instructions=[
                    "Confirme que o menu Caixas/Bancos estÃ¡ visÃ­vel.",
                    "Se quiser testar um seletor, use: x <xpath>.",
                    "Pressione ENTER para o Selenium abrir Caixas/Bancos.",
                ],
            )
            self._click_any(self.MENU_CAIXAS_BANCOS_CANDIDATES, timeout_seconds=max(60, self.timeout))
            time.sleep(5)
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.MENU.CAIXAS_BANCOS",
                phase="AFTER",
                action="CLICK",
                element_name="MENU_CAIXAS_BANCOS",
                locator=self.MENU_CAIXAS_BANCOS_CANDIDATES[0] if self.MENU_CAIXAS_BANCOS_CANDIDATES else None,
                instructions=[
                    "Confirme que a pÃ¡gina Caixas/Bancos abriu corretamente.",
                    "O prÃ³ximo passo Ã© a auditoria da listagem.",
                    "Pressione ENTER para continuar.",
                ],
            )

    def open_new(self, row: ContaOrdemRow) -> None:
        print(f"\n[TRANSFERÃŠNCIA] Abrindo formulÃ¡rio | linha={row.row_number}")

        self._open_caixas_bancos_menu_v1(row)

        transfer_duplicate_row = self._audit_existing_transfer_before_new_v1(row)
        if transfer_duplicate_row:
            raise TransferenciaDuplicadaError(
                "TransferÃªncia duplicada encontrada na listagem antes de abrir o formulÃ¡rio.",
                matching_row=transfer_duplicate_row,
            )

        with step(log, "transfer.open_new", row=row.row_number):
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.NOVA_TRANSFERENCIA",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_NOVA_TRANSFERENCIA",
                locator=self.BTN_NOVA_TRANSFERENCIA_CANDIDATES[0] if self.BTN_NOVA_TRANSFERENCIA_CANDIDATES else None,
                instructions=[
                    "Confirme que o botÃ£o Nova TransferÃªncia estÃ¡ visÃ­vel.",
                    "Pressione ENTER para abrir o formulÃ¡rio de transferÃªncia.",
                ],
            )
            self._click_any(self.BTN_NOVA_TRANSFERENCIA_CANDIDATES, timeout_seconds=max(60, self.timeout))
            time.sleep(2)
            self.a.wait_present(self.CAIXA_SAIDA, timeout_seconds=max(60, self.timeout))

            transfer_history_row = self._audit_existing_transfer_in_form_v1(row)
            if transfer_history_row:
                raise TransferenciaDuplicadaError(
                    "Transferência duplicada encontrada no histórico do formulário antes do preenchimento.",
                    matching_row=transfer_history_row,
                )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.NOVA_TRANSFERENCIA",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_NOVA_TRANSFERENCIA",
                locator=self.BTN_NOVA_TRANSFERENCIA_CANDIDATES[0] if self.BTN_NOVA_TRANSFERENCIA_CANDIDATES else None,
                instructions=[
                    "Confirme que o formulÃ¡rio de transferÃªncia abriu.",
                    "O prÃ³ximo passo serÃ¡ preencher Caixa SaÃ­da.",
                    "Pressione ENTER para continuar.",
                ],
            )

    def fill_and_save(self, row: ContaOrdemRow) -> None:
        print(
            f"[TRANSFERÃŠNCIA] Preenchendo | linha={row.row_number} | "
            f"caixa_saida='{row.caixa_saida}' | caixa_entrada='{row.caixa}' | "
            f"valor='{row.importancia}' | data='{row.data_mov}'"
        )

        self._set_debug_context("input_dados")

        with step(log, "transfer.fill", row=row.row_number):
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CAIXA_SAIDA",
                phase="BEFORE",
                action="SELECT",
                element_name="CAIXA_SAIDA",
                locator=self.CAIXA_SAIDA,
                value=row.caixa_saida,
                instructions=[
                    "Confirme que o campo Caixa SaÃ­da estÃ¡ visÃ­vel.",
                    "O valor esperado vem de CONTAORDEM[CAIXA SAIDA].",
                    "Pressione ENTER para selecionar a caixa de saÃ­da.",
                ],
            )
            self._select2_choose_verified_candidates(
                self.CAIXA_SAIDA_CANDIDATES or [self.CAIXA_SAIDA],
                row.caixa_saida,
                row=row,
                field="CAIXA SAÃDA",
            )
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CAIXA_SAIDA",
                phase="AFTER",
                action="SELECT",
                element_name="CAIXA_SAIDA",
                locator=self.CAIXA_SAIDA,
                value=row.caixa_saida,
                instructions=[
                    "Confirme que a caixa de saÃ­da ficou selecionada.",
                    "O prÃ³ximo passo serÃ¡ preencher o valor.",
                    "Pressione ENTER para continuar.",
                ],
            )

            centro_custo = (row.centro_custo or "PADRÃO").strip() or "PADRÃO"
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CENTRO_CUSTO_SAIDA",
                phase="BEFORE",
                action="SELECT",
                element_name="CENTRO_CUSTO_SAIDA",
                locator=self.CENTRO_CUSTO_SAIDA,
                value=centro_custo,
                instructions=[
                    "Confirme que o campo Centro de Custo Saída está visível.",
                    "O valor esperado vem de CONTAORDEM[CENTRO DE CUSTO].",
                    "Pressione ENTER para selecionar o centro de custo da saída.",
                ],
            )
            self._select2_choose_verified_candidates(
                self.CENTRO_CUSTO_SAIDA_CANDIDATES or [self.CENTRO_CUSTO_SAIDA],
                centro_custo,
                row=row,
                field="CENTRO DE CUSTO SAÍDA",
            )
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CENTRO_CUSTO_SAIDA",
                phase="AFTER",
                action="SELECT",
                element_name="CENTRO_CUSTO_SAIDA",
                locator=self.CENTRO_CUSTO_SAIDA,
                value=centro_custo,
                instructions=[
                    "Confirme que o centro de custo da saída ficou selecionado.",
                    "O próximo passo será preencher o valor da saída.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.VALOR",
                phase="BEFORE",
                action="INPUT",
                element_name="VALOR",
                locator=self.VALOR,
                value=row.importancia,
                instructions=[
                    "Confirme que o campo Valor estÃ¡ visÃ­vel.",
                    "O valor esperado Ã© a importÃ¢ncia da linha atual.",
                    "Pressione ENTER para o Selenium preencher o valor.",
                ],
            )
            self.a.type(self.VALOR, format_amount_for_input(row.importancia))
            time.sleep(0.5)
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.VALOR",
                phase="AFTER",
                action="INPUT",
                element_name="VALOR",
                locator=self.VALOR,
                value=row.importancia,
                instructions=[
                    "Confirme que o valor ficou escrito no campo.",
                    "O prÃ³ximo passo serÃ¡ preencher Caixa Entrada.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CAIXA_ENTRADA",
                phase="BEFORE",
                action="SELECT",
                element_name="CAIXA_ENTRADA",
                locator=self.CAIXA_ENTRADA,
                value=row.caixa,
                instructions=[
                    "Confirme que o campo Caixa Entrada estÃ¡ visÃ­vel.",
                    "O valor esperado vem de CONTAORDEM[CAIXA].",
                    "Pressione ENTER para selecionar a caixa de entrada.",
                ],
            )
            self._select2_choose_verified_candidates(
                self.CAIXA_ENTRADA_CANDIDATES or [self.CAIXA_ENTRADA],
                row.caixa,
                row=row,
                field="CAIXA ENTRADA",
            )
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CAIXA_ENTRADA",
                phase="AFTER",
                action="SELECT",
                element_name="CAIXA_ENTRADA",
                locator=self.CAIXA_ENTRADA,
                value=row.caixa,
                instructions=[
                    "Confirme que a caixa de entrada ficou selecionada.",
                    "O prÃ³ximo passo serÃ¡ preencher a data.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CENTRO_CUSTO_ENTRADA",
                phase="BEFORE",
                action="SELECT",
                element_name="CENTRO_CUSTO_ENTRADA",
                locator=self.CENTRO_CUSTO_ENTRADA,
                value=centro_custo,
                instructions=[
                    "Confirme que o campo Centro de Custo Entrada está visível.",
                    "O valor esperado vem de CONTAORDEM[CENTRO DE CUSTO].",
                    "Pressione ENTER para selecionar o centro de custo da entrada.",
                ],
            )
            self._select2_choose_verified_candidates(
                self.CENTRO_CUSTO_ENTRADA_CANDIDATES or [self.CENTRO_CUSTO_ENTRADA],
                centro_custo,
                row=row,
                field="CENTRO DE CUSTO ENTRADA",
            )
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.CENTRO_CUSTO_ENTRADA",
                phase="AFTER",
                action="SELECT",
                element_name="CENTRO_CUSTO_ENTRADA",
                locator=self.CENTRO_CUSTO_ENTRADA,
                value=centro_custo,
                instructions=[
                    "Confirme que o centro de custo da entrada ficou selecionado.",
                    "O próximo passo será preencher o valor da entrada.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.VALOR_ENTRADA",
                phase="BEFORE",
                action="INPUT",
                element_name="VALOR_ENTRADA",
                locator=self.VALOR_ENTRADA,
                value=row.importancia,
                instructions=[
                    "Confirme que o campo Valor de Entrada está visível.",
                    "O valor esperado é a importância da linha atual.",
                    "Pressione ENTER para preencher o valor de entrada.",
                ],
            )
            self.a.type_any_dom(self.VALOR_ENTRADA_CANDIDATES or [self.VALOR_ENTRADA], format_amount_for_input(row.importancia))
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.VALOR_ENTRADA",
                phase="AFTER",
                action="INPUT",
                element_name="VALOR_ENTRADA",
                locator=self.VALOR_ENTRADA,
                value=row.importancia,
                instructions=[
                    "Confirme que o valor de entrada ficou preenchido.",
                    "O próximo passo será preencher a data.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.DATA",
                phase="BEFORE",
                action="INPUT",
                element_name="DATA",
                locator=self.DATA,
                value=row.data_mov,
                instructions=[
                    "Confirme que o campo Data estÃ¡ visÃ­vel.",
                    "O valor esperado vem de CONTAORDEM[DATA MOV.].",
                    "Pressione ENTER para o Selenium preencher a data.",
                ],
            )
            self.a.type(self.DATA, row.data_mov)
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.DATA",
                phase="AFTER",
                action="INPUT",
                element_name="DATA",
                locator=self.DATA,
                value=row.data_mov,
                instructions=[
                    "Confirme que a data ficou preenchida.",
                    "O prÃ³ximo passo serÃ¡ preencher a descriÃ§Ã£o.",
                    "Pressione ENTER para continuar.",
                ],
            )

            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.DESCRICAO",
                phase="BEFORE",
                action="INPUT",
                element_name="DESCRICAO",
                locator=self.DESCRICAO,
                value=row.descricao_soma,
                instructions=[
                    "Confirme que o campo DescriÃ§Ã£o estÃ¡ visÃ­vel.",
                    "O valor esperado vem de CONTAORDEM[DESCRICAO SOMA].",
                    "Pressione ENTER para o Selenium preencher a descriÃ§Ã£o.",
                ],
            )
            self.a.type(self.DESCRICAO, row.descricao_soma, clear=False)
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.DESCRICAO",
                phase="AFTER",
                action="INPUT",
                element_name="DESCRICAO",
                locator=self.DESCRICAO,
                value=row.descricao_soma,
                instructions=[
                    "Confirme que a descriÃ§Ã£o ficou preenchida.",
                    "O prÃ³ximo passo serÃ¡ salvar a transferÃªncia.",
                    "Pressione ENTER para continuar.",
                ],
            )

            time.sleep(2)
            transfer_history_row = self._audit_existing_transfer_in_form_v1(row)
            if transfer_history_row:
                raise TransferenciaDuplicadaError(
                    "Transferência duplicada encontrada no histórico do formulário antes de salvar.",
                    matching_row=transfer_history_row,
                )

        with step(log, "transfer.save", row=row.row_number):
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.SALVAR",
                phase="BEFORE",
                action="CLICK",
                element_name="BTN_SALVAR",
                locator=self.BTN_SALVAR,
                instructions=[
                    "Confirme que todos os campos da transferÃªncia estÃ£o preenchidos.",
                    "Pressione ENTER para salvar a transferÃªncia.",
                ],
            )
            self._dismiss_alerts()
            try:
                self.a.click_any_visible(self.BTN_SALVAR_CANDIDATES or [self.BTN_SALVAR], timeout_seconds=max(30, self.timeout))
            except Exception:
                try:
                    self.a.click_js(self.BTN_SALVAR)
                except Exception:
                    self.a.click(self.BTN_SALVAR)
            self.a.wait_dom_ready(15)
            time.sleep(2)
            self._dismiss_alerts()
            self._debug_checkpoint(
                row=row,
                stage="TRANSFERENCIA.FORM.SALVAR",
                phase="AFTER",
                action="CLICK",
                element_name="BTN_SALVAR",
                locator=self.BTN_SALVAR,
                instructions=[
                    "Confirme que a transferÃªncia foi salva.",
                    "O prÃ³ximo passo serÃ¡ voltar Ã  lista, se o botÃ£o estiver visÃ­vel.",
                    "Pressione ENTER para continuar.",
                ],
            )

        with step(log, "transfer.back", row=row.row_number):
            if self.a.exists(self.BTN_VOLTAR, timeout_seconds=15):
                self._debug_checkpoint(
                    row=row,
                    stage="TRANSFERENCIA.FORM.VOLTAR",
                    phase="BEFORE",
                    action="CLICK",
                    element_name="BTN_VOLTAR",
                    locator=self.BTN_VOLTAR,
                    instructions=[
                        "Confirme que o formulÃ¡rio jÃ¡ foi salvo.",
                        "Pressione ENTER para voltar para a lista.",
                    ],
                )
                self.a.click_js(self.BTN_VOLTAR)
                self.a.wait_dom_ready(15)
                time.sleep(1)
                self._debug_checkpoint(
                    row=row,
                    stage="TRANSFERENCIA.FORM.VOLTAR",
                    phase="AFTER",
                    action="CLICK",
                    element_name="BTN_VOLTAR",
                    locator=self.BTN_VOLTAR,
                    instructions=[
                        "Confirme que voltou para a lista.",
                        "Pressione ENTER para encerrar o passo de debug.",
                    ],
                )

        print(f"[TRANSFERÃŠNCIA] ConcluÃ­da | linha={row.row_number}")

    def run(self, row: ContaOrdemRow) -> str:
        with step(log, "transfer.run", row=row.row_number, tipo=row.tipo.value):
            try:
                self._set_debug_context("input_dados")
                self._debug_checkpoint(
                    row=row,
                    stage="TRANSFERENCIA.RUN.INICIO",
                    phase="BEFORE",
                    action="FLOW",
                    element_name="TRANSFERENCIAS",
                    locator=self.MENU_CAIXAS_BANCOS_CANDIDATES[0] if self.MENU_CAIXAS_BANCOS_CANDIDATES else None,
                    instructions=[
                        "Confirme que o browser estÃ¡ visÃ­vel e logado no SOMA.",
                        "Se o modo selector debug estiver ativo, use x/css para inspecionar seletores.",
                        "Pressione ENTER para iniciar o fluxo de transferÃªncia.",
                    ],
                )
                self.open_new(row)
                self.fill_and_save(row)
                self._debug_checkpoint(
                    row=row,
                    stage="TRANSFERENCIA.RUN.FIM",
                    phase="AFTER",
                    action="FLOW",
                    element_name="TRANSFERENCIAS",
                    locator=self.MENU_CAIXAS_BANCOS_CANDIDATES[0] if self.MENU_CAIXAS_BANCOS_CANDIDATES else None,
                    instructions=[
                        "Confirme que a transferÃªncia terminou sem erros.",
                        "Pressione ENTER para finalizar o debug deste registo.",
                    ],
                )
                return "Transferido"
            finally:
                self._set_debug_context(None)
