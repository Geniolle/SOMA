from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.config.fields import FORM_FIELD_REGISTRY, field_names


class FakeActions:
    def __init__(self, present, *, visible=None, enabled=None, texts=None, close_on_click=None):
        self.present = set(present)
        self.visible = set(visible if visible is not None else present)
        self.enabled = set(enabled if enabled is not None else present)
        self.texts = dict(texts or {})
        self.close_on_click = set(close_on_click or [])
        self.selected = []
        self.screenshots = []
        self.typed = []
        self.clicked = []
        self.js_clicked = []
        self.wait_invisible_calls = []
        self.wait_any_present_calls = 0
        self.driver = SimpleNamespace(
            current_url="https://test.local/",
            title="SOMA",
            execute_script=lambda *args, **kwargs: None,
            switch_to=SimpleNamespace(alert=property(lambda self: None)),
        )

    def exists(self, locator, timeout_seconds=None):
        return locator in self.present

    def wait_any_present(self, locators, timeout_seconds=None):
        self.wait_any_present_calls += 1
        for loc in locators:
            if loc in self.present:
                return loc
        raise TimeoutException("missing")

    def wait_visible(self, locator, timeout_seconds=None):
        if locator in self.present and locator in self.visible:
            return SimpleNamespace(
                text=self.texts.get(locator, ""),
                is_displayed=lambda: True,
                is_enabled=lambda: locator in self.enabled,
            )
        raise TimeoutException("not visible")

    def wait_clickable(self, locator, timeout_seconds=None):
        if locator in self.present and locator in self.visible and locator in self.enabled:
            return SimpleNamespace(
                text=self.texts.get(locator, ""),
                is_displayed=lambda: True,
                is_enabled=lambda: True,
            )
        raise TimeoutException("not clickable")

    def wait_invisible(self, locator, timeout_seconds=None):
        self.wait_invisible_calls.append(locator)
        if locator in self.present and locator in self.visible:
            raise TimeoutException("still visible")
        return None

    def select2_choose(self, locator, value):
        self.selected.append((locator, value))

    def type(self, locator, value, clear=True):
        self.typed.append((locator, value, clear))

    def click_js(self, locator):
        self.js_clicked.append(locator)
        for item in self.close_on_click:
            self.present.discard(item)
            self.visible.discard(item)
        return None

    def click(self, locator):
        self.clicked.append(locator)
        for item in self.close_on_click:
            self.present.discard(item)
            self.visible.discard(item)
        return None

    def screenshot(self, name):
        self.screenshots.append(name)
        return Path(f"{name}.png")


def _build_page(actions, settings=None):
    settings = settings or SimpleNamespace(site_home_url="", STRICT_CAIXA_MATCH=True)
    return EntradasSaidasPage(actions, settings)


def test_select2_choose_candidates_uses_first_available_locator():
    missing = ("xpath", "//missing")
    fallback = ("xpath", "//fallback")
    actions = FakeActions({fallback})
    page = _build_page(actions)
    row = SimpleNamespace(row_number=3375)

    page._select2_choose_candidates([missing, fallback], "Plano Teste", row=row, field="plano_conta")

    assert actions.selected == [(fallback, "Plano Teste")]


def test_select2_choose_candidates_raises_when_no_locator_is_present():
    actions = FakeActions(set())
    page = _build_page(actions)
    row = SimpleNamespace(row_number=3375)

    try:
        page._select2_choose_candidates([("xpath", "//missing")], "Plano Teste", row=row, field="plano_conta")
    except TimeoutException as exc:
        assert "opener do Select2" in str(exc)
    else:
        raise AssertionError("Era esperado TimeoutException quando nenhum locator esta presente.")


def test_type_and_validate_candidates_returns_actual_field_value(monkeypatch):
    locator = ("xpath", "//descricao")
    actions = FakeActions({locator})
    page = _build_page(actions)
    row = SimpleNamespace(row_number=3375)
    monkeypatch.setattr(page, "_input_value", lambda loc: "VENDA DA CANTINA (VERBO CAFE) N001")

    value = page._type_and_validate_candidates(
        [locator],
        "VENDA DA CANTINA (VERBO CAFE) N001",
        row=row,
        field="descricao",
    )

    assert value == "VENDA DA CANTINA (VERBO CAFE) N001"
    assert actions.typed == [(locator, "VENDA DA CANTINA (VERBO CAFE) N001", True)]


def test_type_and_validate_candidates_raises_when_field_stays_empty(monkeypatch):
    locator = ("xpath", "//descricao")
    actions = FakeActions({locator})
    page = _build_page(actions)
    row = SimpleNamespace(row_number=3375)
    monkeypatch.setattr(page, "_input_value", lambda loc: "")

    try:
        page._type_and_validate_candidates(
            [locator],
            "VENDA DA CANTINA (VERBO CAFE) N001",
            row=row,
            field="descricao",
        )
    except RuntimeError as exc:
        assert "confirmou preenchimento" in str(exc)
    else:
        raise AssertionError("Era esperado RuntimeError quando o campo continua vazio apos digitacao.")


def test_page_loads_description_locator_from_external_file(tmp_path):
    locators_path = tmp_path / "locators.json"
    from soma_app.config.locators import _is_locator, _is_locator_list
    registry_fields = set(field_names(FORM_FIELD_REGISTRY.entradas_saidas.common))
    registry_fields.update(field_names(FORM_FIELD_REGISTRY.entradas_saidas.entrada_specific))
    registry_fields.update(field_names(FORM_FIELD_REGISTRY.entradas_saidas.saida_specific))
    registry_fields.update(field_names(FORM_FIELD_REGISTRY.entradas_saidas.search))
    registry_fields.update(field_names(FORM_FIELD_REGISTRY.transferencias))
    registry_fields.update(field_names(FORM_FIELD_REGISTRY.caixas_bancos))

    entradas_saidas_cfg = {}
    for name, val in vars(EntradasSaidasPage).items():
        if name.isupper() and name not in {"RADIO_ANY_CANDIDATES", "FORM_READY_CANDIDATES", "ANY_VALUE_CANDIDATES"}:
            if _is_locator(val):
                entradas_saidas_cfg[name] = f"//dummy-{name.lower()}"
            elif _is_locator_list(val):
                entradas_saidas_cfg[name] = [f"//dummy-{name.lower()}"]

    for field_name_value in registry_fields:
        entradas_saidas_cfg.setdefault(field_name_value, f"//dummy-{field_name_value.lower()}")
    entradas_saidas_cfg["DESCRICAO"] = "//input[@name='descricao_custom']"
    entradas_saidas_cfg["DESCRICAO_CANDIDATES"] = [
        "//input[@name='descricao_custom']",
        "//dummy-desc-cand-1"
    ]

    locators_path.write_text(
        json.dumps(
            {
                "entradas_saidas": entradas_saidas_cfg
            }
        ),
        encoding="utf-8",
    )
    settings = SimpleNamespace(
        site_home_url="",
        STRICT_CAIXA_MATCH=True,
        locators_path=locators_path,
    )

    page = _build_page(FakeActions(set()), settings=settings)

    assert page.DESCRICAO == ("xpath", "//input[@name='descricao_custom']")
    assert page.DESCRICAO_CANDIDATES[0] == page.DESCRICAO
    assert page.FORM_READY_CANDIDATES[1] == page.DESCRICAO


def test_confirmation_popup_clicks_sim_and_closes(monkeypatch):
    popup_button = ("xpath", "//div[contains(@class,'swal2-container')]//button[normalize-space()='Sim']")
    popup_container = ("xpath", "//div[contains(@class,'swal2-popup') and contains(normalize-space(.), 'realizar o pagamento')]")
    actions = FakeActions(
        {popup_button, popup_container},
        visible={popup_button, popup_container},
        enabled={popup_button, popup_container},
        texts={popup_button: "Sim"},
        close_on_click={popup_button, popup_container},
    )
    page = _build_page(actions)
    row = SimpleNamespace(row_number=22, tipo=SimpleNamespace(value="Entrada"))

    handled = page._handle_confirmation_popup(row=row, accept=True, timeout_seconds=1)

    assert handled is True
    assert actions.clicked == [popup_button]
    assert popup_button in actions.wait_invisible_calls or popup_container in actions.wait_invisible_calls


def test_save_form_if_present_raises_when_confirmation_popup_stays_open(monkeypatch):
    save_button = ("xpath", "//button[@id='btn_salvar']")
    popup_container = (By.CLASS_NAME, "swal2-container")
    actions = FakeActions({save_button, popup_container}, visible={save_button, popup_container}, enabled={save_button, popup_container})
    page = _build_page(actions)
    page.BTN_SALVAR_FORM = save_button
    row = SimpleNamespace(row_number=22, tipo=SimpleNamespace(value="Entrada"), forma_pagamento="TRANSFERENCIA BANCARIA")

    monkeypatch.setattr(page, "_dismiss_overlays_with_wait", lambda max_wait_seconds=3: None)
    monkeypatch.setattr(page, "_close_datepicker", lambda: None)
    monkeypatch.setattr(page, "_handle_confirmation_popup", lambda **kwargs: False)

    try:
        page._save_form_if_present(row)
    except RuntimeError as exc:
        assert "Popup de confirmacao permaneceu aberto" in str(exc)
    else:
        raise AssertionError("Era esperado RuntimeError quando o popup nao fosse tratado.")


def test_transferencia_bancaria_normalization_is_tolerant():
    page = _build_page(FakeActions(set()))

    assert page._is_transferencia_bancaria("TRANSFERENCIA BANCARIA")
    assert page._is_transferencia_bancaria(" Transferência Bancária ")
    assert page._is_transferencia_bancaria("transferencia bancaria")
    assert not page._is_transferencia_bancaria(None)


def test_exists_any_uses_single_aggregate_wait():
    missing = ("xpath", "//missing")
    found = ("xpath", "//found")
    actions = FakeActions({found})
    page = _build_page(actions)

    assert page._exists_any([missing, found], timeout_seconds=1) is True
    assert actions.wait_any_present_calls == 1


def test_create_entry_skips_separate_payment_modal(monkeypatch):
    actions = FakeActions(set())
    page = _build_page(actions)
    row = SimpleNamespace(
        row_number=10,
        tipo=SimpleNamespace(value="Entrada"),
        forma_pagamento="TRANSFERENCIA BANCARIA",
        data_mov="29/07/2026",
    )

    called = []

    monkeypatch.setattr(page, "_open_new", lambda row: called.append("open_new"))
    monkeypatch.setattr(page, "_choose_tipo", lambda row: called.append("choose_tipo"))
    monkeypatch.setattr(page, "_fill_common", lambda row: called.append("fill_common"))
    monkeypatch.setattr(page, "_fill_entrada_sem_caixa", lambda row: called.append("fill_entrada"))
    monkeypatch.setattr(page, "_fill_caixa_entrada_ultima", lambda row: called.append("fill_caixa"))
    monkeypatch.setattr(page, "_save_form_if_present", lambda row: called.append("save"))
    monkeypatch.setattr(page, "_realizar_pagamento", lambda row: called.append("realizar_pagamento"))
    monkeypatch.setattr(page, "_pagamento_saida_modal", lambda row: called.append("pagamento_saida_modal"))
    monkeypatch.setattr(page, "_do_baixa", lambda row: called.append("baixa"))
    monkeypatch.setattr(page, "_search_doc_id", lambda row: "DOC-1")
    monkeypatch.setattr(page, "_confirm", lambda *args, **kwargs: None)

    doc = page.create_and_get_doc_id(row)

    assert doc == "DOC-1"
    assert "realizar_pagamento" not in called
    assert "pagamento_saida_modal" not in called
    assert "baixa" in called
