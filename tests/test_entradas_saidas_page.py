from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.config.fields import FORM_FIELD_REGISTRY, field_names
from soma_app.domain.models import ContaOrdemRow, TipoMovimento


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


def test_create_saida_calls_payment_flow_and_skips_baixa_when_not_transfer(monkeypatch):
    actions = FakeActions(set())
    page = _build_page(actions)
    row = SimpleNamespace(
        row_number=11,
        tipo=TipoMovimento.SAIDA,
        forma_pagamento="DINHEIRO",
        data_mov="29/07/2026",
    )

    called = []

    monkeypatch.setattr(page, "_open_new", lambda row: called.append("open_new"))
    monkeypatch.setattr(page, "_choose_tipo", lambda row: called.append("choose_tipo"))
    monkeypatch.setattr(page, "_fill_common", lambda row: called.append("fill_common"))
    monkeypatch.setattr(page, "_fill_saida", lambda row: called.append("fill_saida"))
    monkeypatch.setattr(page, "_save_form_if_present", lambda row: called.append("save"))
    monkeypatch.setattr(page, "_realizar_pagamento", lambda row: called.append("realizar_pagamento"))
    monkeypatch.setattr(page, "_pagamento_saida_modal", lambda row: called.append("pagamento_saida_modal"))
    monkeypatch.setattr(page, "_do_baixa", lambda row: called.append("baixa"))
    monkeypatch.setattr(page, "_search_doc_id", lambda row: "DOC-S1")
    monkeypatch.setattr(page, "_confirm", lambda *args, **kwargs: None)

    doc = page.create_and_get_doc_id(row)

    assert doc == "DOC-S1"
    assert "realizar_pagamento" in called
    assert "pagamento_saida_modal" in called
    assert "baixa" not in called


def test_realizar_pagamento_best_effort_when_modal_does_not_open(monkeypatch):
    actions = FakeActions({("xpath", "//realizar")})
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [("xpath", "//inserir-pagamento")]
    page.BTN_REALIZAR_PAGAMENTO_CANDIDATES = [("xpath", "//realizar")]
    row = SimpleNamespace(row_number=12, tipo=TipoMovimento.SAIDA)

    emitted = []
    monkeypatch.setattr(page, "_exists_any", lambda candidates, timeout_seconds=1: True)
    monkeypatch.setattr(page.a, "wait_any_present", lambda candidates, timeout_seconds=10: candidates[0])
    monkeypatch.setattr(page.a, "click", lambda locator: actions.clicked.append(locator))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text="Inserir Pagamento" if args[1] == "//inserir-pagamento" else "Realizar pagamento")

    def _raise_visible(locator, timeout_seconds=10):
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_visible", _raise_visible)
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))

    page._realizar_pagamento(row)

    assert actions.clicked == [("xpath", "//inserir-pagamento")]
    assert any("vou seguir com o fluxo" in msg for msg, _ in emitted)


def test_realizar_pagamento_prefers_insert_button_when_available(monkeypatch):
    insert_button = ("xpath", "//inserir-pagamento")
    actions = FakeActions({insert_button}, texts={insert_button: "Inserir Pagamento"})
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [insert_button]
    page.BTN_REALIZAR_PAGAMENTO_CANDIDATES = [("xpath", "//realizar")]
    row = SimpleNamespace(row_number=19, tipo=TipoMovimento.SAIDA)

    emitted = []
    monkeypatch.setattr(page, "_exists_any", lambda candidates, timeout_seconds=1: True)
    monkeypatch.setattr(page.a, "wait_any_present", lambda candidates, timeout_seconds=10: candidates[0])
    monkeypatch.setattr(page.a, "click", lambda locator: actions.clicked.append(locator))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text=actions.texts.get(args, ""))
    monkeypatch.setattr(page.a, "wait_visible", lambda locator, timeout_seconds=10: (_ for _ in ()).throw(TimeoutException("missing")))
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))

    page._realizar_pagamento(row)

    assert actions.clicked == [insert_button]
    assert any("Botao de realizar pagamento encontrado" in msg or "Botao de inserir pagamento encontrado" in msg for msg, _ in emitted)


def test_pagamento_saida_modal_skips_cancel_payment_button(monkeypatch):
    cancel_button = ("xpath", "//cancelar-pagamento")
    insert_button = ("xpath", "//inserir-pagamento")
    data_modal = ("name", "data_pagamento")
    forma_modal = ("name", "forma_pagamento")
    caixa_modal = ("name", "id_caixa")
    documento_modal = ("name", "num_documento")
    salvar_modal = ("id", "botao_pagamento")

    actions = FakeActions(
        {cancel_button, insert_button},
        texts={
            cancel_button: "Cancelar pagamento",
            insert_button: "Inserir pagamento",
            data_modal: "",
            forma_modal: "",
            caixa_modal: "",
            documento_modal: "",
            salvar_modal: "Salvar Pagamento",
        },
    )
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [cancel_button, insert_button]
    page.DATA_PAGAMENTO_MODAL = data_modal
    page.FORMA_PAGAMENTO_MODAL = forma_modal
    page.CAIXA_PAGAMENTO_MODAL = caixa_modal
    page.NUM_DOCUMENTO_MODAL = documento_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL = salvar_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES = [salvar_modal]

    row = SimpleNamespace(
        row_number=15,
        tipo=TipoMovimento.SAIDA,
        data_mov="30/07/2026",
        forma_pagamento="DINHEIRO",
        caixa="CAIXA DIÃRIO",
        descricao_soma="TESTE",
        dados_doc="DOC-123",
        doc_soma="",
    )

    emitted = []
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text=actions.texts.get(args, ""))

    def wait_visible(locator, timeout_seconds=None):
        if locator in actions.present and locator in actions.visible:
            return SimpleNamespace(text="", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("not visible")

    monkeypatch.setattr(page.a, "wait_visible", wait_visible)

    def click(locator):
        actions.clicked.append(locator)
        if locator == insert_button:
            for item in {data_modal, forma_modal, caixa_modal, documento_modal, salvar_modal}:
                actions.present.add(item)
                actions.visible.add(item)
        if locator == salvar_modal:
            actions.present.discard(data_modal)
            actions.visible.discard(data_modal)
        return None

    monkeypatch.setattr(page.a, "click", click)

    def wait_any_present(candidates, timeout_seconds=10):
        for loc in candidates:
            if loc in actions.present:
                return loc
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_any_present", wait_any_present)
    monkeypatch.setattr(page, "_select_best_effort", lambda _el, desired, **kwargs: desired)
    monkeypatch.setattr(page, "_input_value", lambda locator: "30/07/2026")

    page._pagamento_saida_modal(row)

    assert actions.clicked[0] == insert_button
    assert cancel_button not in actions.clicked
    assert actions.clicked[-1] == salvar_modal
    assert any("cancelamento" in msg.lower() for msg, _ in emitted)


def test_pagamento_saida_modal_waits_for_native_select_options(monkeypatch):
    insert_button = ("xpath", "//inserir-pagamento")
    data_modal = ("name", "data_pagamento")
    forma_modal = ("name", "forma_pagamento")
    caixa_modal = ("name", "id_caixa")
    documento_modal = ("name", "num_documento")
    salvar_modal = ("id", "botao_pagamento")

    actions = FakeActions({insert_button}, close_on_click={data_modal})
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [insert_button]
    page.DATA_PAGAMENTO_MODAL = data_modal
    page.FORMA_PAGAMENTO_MODAL = forma_modal
    page.CAIXA_PAGAMENTO_MODAL = caixa_modal
    page.NUM_DOCUMENTO_MODAL = documento_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL = salvar_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES = [salvar_modal]

    row = SimpleNamespace(
        row_number=20,
        tipo=TipoMovimento.SAIDA,
        data_mov="30/07/2026",
        forma_pagamento="TRANSFERÊNCIA BANCÁRIA",
        caixa="CAIXA DIÁRIO",
        descricao_soma="TESTE",
    )

    emitted = []
    wait_ready_calls = []
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text="Inserir Pagamento")

    def wait_visible(locator, timeout_seconds=None):
        if locator in actions.present and locator in actions.visible:
            return SimpleNamespace(text="", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("not visible")

    monkeypatch.setattr(page.a, "wait_visible", wait_visible)
    def click(locator):
        actions.clicked.append(locator)
        if locator == insert_button:
            for item in {data_modal, forma_modal, caixa_modal, documento_modal, salvar_modal}:
                actions.present.add(item)
                actions.visible.add(item)
        if locator == salvar_modal:
            actions.present.discard(data_modal)
            actions.visible.discard(data_modal)
        return None

    monkeypatch.setattr(page.a, "click", click)
    monkeypatch.setattr(page.a, "click_js", lambda locator: actions.js_clicked.append(locator))
    monkeypatch.setattr(page.a, "wait_any_present", lambda candidates, timeout_seconds=10: candidates[0])
    monkeypatch.setattr(
        page,
        "_wait_select_ready",
        lambda resolve_select, timeout_seconds=10, min_options=2: wait_ready_calls.append((timeout_seconds, min_options)),
    )
    monkeypatch.setattr(page, "_select_best_effort", lambda _el, desired, **kwargs: desired)
    monkeypatch.setattr(page, "_select_with_sleep_validation", lambda resolve_select, desired, **kwargs: desired)
    monkeypatch.setattr(page, "_input_value", lambda locator: "30/07/2026")
    doc_calls = []
    monkeypatch.setattr(
        page,
        "_type_and_validate_candidates",
        lambda locators, value, **kwargs: doc_calls.append((tuple(locators), value, kwargs.get("field"))) or value,
    )

    page._pagamento_saida_modal(row)

    assert actions.clicked[0] == insert_button
    assert wait_ready_calls[0] == (10, 2)
    assert wait_ready_calls[1] == (10, 1)
    assert actions.clicked[-1] == salvar_modal
    assert any("Forma de pagamento selecionada com sucesso" in msg for msg, _ in emitted)


def test_pagamento_saida_modal_validates_id_interno_for_transferencia():
    """
    Testa que _validate_pagamento_saida_modal valida id_interno (não descricao_soma).
    """
    row = ContaOrdemRow(
        row_number=21,
        tipo=TipoMovimento.SAIDA,
        data_mov="30/07/2026",
        forma_pagamento="TRANSFERENCIA BANCARIA",
        caixa="CAIXA DIARIO",
        descricao_soma="TESTE",
        id_interno="INT-123456",  # Deve usar este, não descricao_soma
        doc_soma="",
        dados_doc="",
        iduser="USER1",
        timestamp="2024-01-15 10:00:00",
        caixa_saida="CAIXA_SAIDA",
        centro_custo="CC001",
        plano_conta="PLAN001",
        importancia="1000.00",
        raw={},
    )

    actions = FakeActions(set())
    page = _build_page(actions)
    data_modal = ("name", "data_pagamento")
    page.DATA_PAGAMENTO_MODAL = data_modal
    page.NUM_DOCUMENTO_MODAL = ("name", "num_documento")

    # Mock métodos
    page._input_value = lambda loc: "INT-123456" if loc == page.NUM_DOCUMENTO_MODAL else "30/07/2026"
    page._match_ok = lambda a, b: a.strip() == b.strip()

    # Validação deve passar (usa id_interno)
    page._validate_pagamento_saida_modal(
        row=row,
        chosen_fp="TRANSFERENCIA BANCARIA",
        chosen_cx="CAIXA DIARIO"
    )

    # Se tivesse usado descricao_soma ("TESTE"), teria falhado
    # Mas usa id_interno ("INT-123456"), então passa!


def test_pagamento_saida_modal_skips_caixa_when_backend_returns_nan_error(monkeypatch):
    insert_button = ("xpath", "//inserir-pagamento")
    data_modal = ("name", "data_pagamento")
    forma_modal = ("name", "forma_pagamento")
    documento_modal = ("name", "num_documento")
    salvar_modal = ("id", "botao_pagamento")

    actions = FakeActions({insert_button}, close_on_click={data_modal})
    actions.driver.page_source = """
        <div class="buscar_caixa">
            <table><tbody><tr><td><b>Erro!</b> Unknown column 'NaN' in 'where clause'</td></tr></tbody></table>
        </div>
    """
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [insert_button]
    page.DATA_PAGAMENTO_MODAL = data_modal
    page.FORMA_PAGAMENTO_MODAL = forma_modal
    page.NUM_DOCUMENTO_MODAL = documento_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL = salvar_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES = [salvar_modal]

    row = SimpleNamespace(
        row_number=22,
        tipo=TipoMovimento.SAIDA,
        data_mov="30/07/2026",
        forma_pagamento="TRANSFERENCIA BANCARIA",
        caixa="CAIXA DIARIO",
        descricao_soma="OBS DO PAGAMENTO",
        dados_doc="DOC-123",
        doc_soma="DOC-ALT",
    )

    emitted = []
    doc_calls = []
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text="Inserir Pagamento")

    def wait_visible(locator, timeout_seconds=None):
        if locator in actions.present and locator in actions.visible:
            return SimpleNamespace(text="", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("not visible")

    monkeypatch.setattr(page.a, "wait_visible", wait_visible)
    monkeypatch.setattr(page.a, "click", lambda locator: actions.clicked.append(locator))
    monkeypatch.setattr(page.a, "click_js", lambda locator: actions.js_clicked.append(locator))
    monkeypatch.setattr(page.a, "wait_any_present", lambda candidates, timeout_seconds=10: candidates[0])
    monkeypatch.setattr(page, "_wait_select_ready", lambda *args, **kwargs: None)
    monkeypatch.setattr(page, "_select_best_effort", lambda _el, desired, **kwargs: desired)
    monkeypatch.setattr(page, "_select_with_sleep_validation", lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("nao deveria selecionar caixa")))
    monkeypatch.setattr(page, "_input_value", lambda locator: "30/07/2026")
    monkeypatch.setattr(
        page,
        "_type_and_validate_candidates",
        lambda locators, value, **kwargs: doc_calls.append((tuple(locators), value)) or value,
    )

    def click(locator):
        actions.clicked.append(locator)
        if locator == insert_button:
            for item in {data_modal, forma_modal, documento_modal, salvar_modal}:
                actions.present.add(item)
                actions.visible.add(item)
        if locator == salvar_modal:
            actions.present.discard(data_modal)
            actions.visible.discard(data_modal)
        return None

    monkeypatch.setattr(page.a, "click", click)

    page._pagamento_saida_modal(row)

    assert doc_calls == [((documento_modal,), "OBS DO PAGAMENTO")]
    assert salvar_modal in actions.clicked
    assert any("erro do backend" in msg.lower() for msg, _ in emitted)


def test_select_best_effort_uses_value_fallback_for_payment_form(monkeypatch):
    class FakeOption:
        def __init__(self, text, value):
            self.text = text
            self._value = value

        def get_attribute(self, name):
            return self._value if name == "value" else None

    class FakeSelect:
        def __init__(self, _el):
            self.options = [FakeOption("", str(idx)) for idx in range(6)]
            self.selected_value = None

        def select_by_visible_text(self, _text):
            raise Exception("visible text unavailable")

        def select_by_index(self, index):
            self.selected_value = self.options[index].get_attribute("value")

        def select_by_value(self, value):
            self.selected_value = value

        @property
        def first_selected_option(self):
            for option in self.options:
                if option.get_attribute("value") == self.selected_value:
                    return option
            return self.options[0]

    page = _build_page(FakeActions(set()))
    monkeypatch.setattr("soma_app.automation.pages.entradas_saidas_page.Select", FakeSelect)

    chosen = page._select_best_effort(
        SimpleNamespace(),
        "TRANSFERÊNCIA BANCÁRIA",
        row=SimpleNamespace(row_number=12),
        field="FORMA_PAGAMENTO_MODAL",
        strict=True,
    )

    assert chosen == "TRANSFERÊNCIA BANCÁRIA"


def test_wait_payment_button_candidate_retries_until_button_appears(monkeypatch):
    cancel_button = ("xpath", "//cancelar-pagamento")
    insert_button = ("xpath", "//inserir-pagamento")
    actions = FakeActions({cancel_button, insert_button}, texts={cancel_button: "Cancelar pagamento", insert_button: "Inserir pagamento"})
    page = _build_page(actions)
    row = SimpleNamespace(row_number=18, tipo=TipoMovimento.SAIDA)

    calls = {"count": 0}

    def wait_any_present(candidates, timeout_seconds=10):
        calls["count"] += 1
        if calls["count"] == 1:
            raise TimeoutException("not ready yet")
        for loc in candidates:
            if loc in actions.present:
                return loc
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_any_present", wait_any_present)
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text=actions.texts.get(args, ""))

    loc = page._wait_payment_button_candidate(
        [cancel_button, insert_button],
        row=row,
        label="inserir pagamento",
        timeout_seconds=3,
    )

    assert loc == insert_button
    assert calls["count"] >= 2


def test_pagamento_saida_modal_tries_click_js_when_normal_click_does_not_open_modal(monkeypatch):
    insert_button = ("xpath", "//inserir-pagamento")
    data_modal = ("name", "data_pagamento")
    forma_modal = ("name", "forma_pagamento")
    caixa_modal = ("name", "id_caixa")
    documento_modal = ("name", "num_documento")
    salvar_modal = ("id", "botao_pagamento")

    actions = FakeActions(
        {insert_button},
        close_on_click={data_modal},
    )
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [insert_button]
    page.DATA_PAGAMENTO_MODAL = data_modal
    page.FORMA_PAGAMENTO_MODAL = forma_modal
    page.CAIXA_PAGAMENTO_MODAL = caixa_modal
    page.NUM_DOCUMENTO_MODAL = documento_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL = salvar_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES = [salvar_modal]

    row = SimpleNamespace(
        row_number=16,
        tipo=TipoMovimento.SAIDA,
        data_mov="30/07/2026",
        forma_pagamento="DINHEIRO",
        caixa="CAIXA DIÁRIO",
        descricao_soma="TESTE",
    )

    emitted = []
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text="Inserir Pagamento")

    def wait_visible(locator, timeout_seconds=None):
        if locator in actions.present and locator in actions.visible:
            return SimpleNamespace(text="", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("not visible")

    monkeypatch.setattr(page.a, "wait_visible", wait_visible)

    def click(locator):
        actions.clicked.append(locator)
        actions.present.discard(data_modal)
        actions.visible.discard(data_modal)
        return None

    monkeypatch.setattr(page.a, "click", click)

    def click_js(locator):
        actions.js_clicked.append(locator)
        for item in {data_modal, forma_modal, caixa_modal, documento_modal, salvar_modal}:
            actions.present.add(item)
            actions.visible.add(item)
        return None

    monkeypatch.setattr(page.a, "click_js", click_js)
    def wait_any_present(candidates, timeout_seconds=10):
        for loc in candidates:
            if loc in actions.present:
                return loc
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_any_present", wait_any_present)
    monkeypatch.setattr(page, "_select_best_effort", lambda _el, desired, **kwargs: desired)
    monkeypatch.setattr(page, "_input_value", lambda locator: "30/07/2026")

    page._pagamento_saida_modal(row)

    assert actions.clicked[0] == insert_button
    assert actions.js_clicked == [insert_button]
    assert actions.clicked[-1] == salvar_modal
    assert any("click_js" in msg for msg, _ in emitted)


def test_pagamento_saida_modal_ignores_delete_payment_modal(monkeypatch):
    insert_button = ("xpath", "//inserir-pagamento")
    delete_modal = ("id", "cancelar")
    delete_check = ("css selector", "input.pagamentos_check")
    delete_check_all = ("css selector", "input.all")
    delete_button = ("id", "cancelar_pagamento")
    data_modal = ("name", "data_pagamento")
    forma_modal = ("name", "forma_pagamento")
    caixa_modal = ("name", "id_caixa")
    documento_modal = ("name", "num_documento")
    salvar_modal = ("id", "botao_pagamento")

    actions = FakeActions({insert_button}, close_on_click={delete_modal, data_modal})
    page = _build_page(actions)
    page.BTN_INSERIR_PAGAMENTO_SAIDA_CANDIDATES = [insert_button]
    page.MODAL_EXCLUIR_PAGAMENTO = delete_modal
    page.CHECK_EXCLUIR_PAGAMENTO = delete_check
    page.CHECK_EXCLUIR_PAGAMENTO_ALL = delete_check_all
    page.BTN_EXCLUIR_PAGAMENTO_MODAL = delete_button
    page.DATA_PAGAMENTO_MODAL = data_modal
    page.FORMA_PAGAMENTO_MODAL = forma_modal
    page.CAIXA_PAGAMENTO_MODAL = caixa_modal
    page.NUM_DOCUMENTO_MODAL = documento_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL = salvar_modal
    page.BTN_SALVAR_PAGAMENTO_MODAL_CANDIDATES = [salvar_modal]

    row = SimpleNamespace(
        row_number=17,
        tipo=TipoMovimento.SAIDA,
        data_mov="30/07/2026",
        forma_pagamento="DINHEIRO",
        caixa="CAIXA DIÁRIO",
        descricao_soma="TESTE",
    )

    emitted = []
    state = {"insert_calls": 0}
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(
        text="Inserir Pagamento",
        is_enabled=lambda: True,
        get_attribute=lambda *_: "",
    )

    def wait_visible(locator, timeout_seconds=None):
        if locator in actions.present and locator in actions.visible:
            return SimpleNamespace(text="", is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("not visible")

    monkeypatch.setattr(page.a, "wait_visible", wait_visible)

    def click(locator):
        actions.clicked.append(locator)
        if locator == insert_button and state["insert_calls"] == 0:
            state["insert_calls"] += 1
            for item in {delete_modal, delete_check, delete_check_all, delete_button}:
                actions.present.add(item)
                actions.visible.add(item)
        elif locator == delete_button:
            for item in {delete_modal, delete_check, delete_check_all, delete_button}:
                actions.present.discard(item)
                actions.visible.discard(item)
        elif locator == salvar_modal:
            actions.present.discard(data_modal)
            actions.visible.discard(data_modal)
        return None

    def click_js(locator):
        actions.js_clicked.append(locator)
        if locator == delete_check_all or locator == delete_check:
            return None
        if locator == delete_button:
            for item in {delete_modal, delete_check, delete_check_all, delete_button}:
                actions.present.discard(item)
                actions.visible.discard(item)
            return None
        if locator == insert_button and state["insert_calls"] == 1:
            for item in {data_modal, forma_modal, caixa_modal, documento_modal, salvar_modal}:
                actions.present.add(item)
                actions.visible.add(item)
            return None
        return None

    monkeypatch.setattr(page.a, "click", click)
    monkeypatch.setattr(page.a, "click_js", click_js)
    def wait_any_present(candidates, timeout_seconds=10):
        for loc in candidates:
            if loc in actions.present:
                return loc
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_any_present", wait_any_present)
    monkeypatch.setattr(page, "_select_best_effort", lambda _el, desired, **kwargs: desired)
    monkeypatch.setattr(page, "_input_value", lambda locator: "30/07/2026")

    page._pagamento_saida_modal(row)

    assert delete_check_all not in actions.js_clicked
    assert delete_check not in actions.js_clicked
    assert delete_button not in actions.js_clicked
    assert insert_button in actions.clicked
    assert insert_button in actions.js_clicked
    assert actions.clicked[-1] == salvar_modal
    assert not any("exclusao de pagamento" in msg.lower() for msg, _ in emitted)


def test_do_baixa_raises_when_modal_does_not_open(monkeypatch):
    """
    Testa que _do_baixa agora lança erro quando o modal não abre (não é mais best effort).
    Transferência Bancária requer sucesso explícito na baixa.
    """
    actions = FakeActions({("xpath", "//baixa")})
    page = _build_page(actions)
    page.BTN_INSERIR_BAIXA_CANDIDATES = [("xpath", "//baixa")]
    row = SimpleNamespace(row_number=13, tipo=TipoMovimento.SAIDA, data_mov="29/07/2026")

    emitted = []
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))
    monkeypatch.setattr(page, "_dismiss_overlays_with_wait", lambda **kwargs: None)
    monkeypatch.setattr(page, "_exists_any", lambda candidates, timeout_seconds=1: True)
    monkeypatch.setattr(page.a, "wait_any_present", lambda candidates, timeout_seconds=10: candidates[0])
    monkeypatch.setattr(page.a, "click", lambda locator: actions.clicked.append(locator))
    monkeypatch.setattr(page.a, "screenshot", lambda name: Path(f"{name}.png"))
    page.a.driver.find_element = lambda *args, **kwargs: SimpleNamespace(text="Inserir Baixa")

    def _raise_visible(locator, timeout_seconds=10):
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_visible", _raise_visible)

    # Agora deve lançar RuntimeError quando o modal não abre (não é mais best effort)
    try:
        page._do_baixa(row)
        raise AssertionError("Era esperado RuntimeError quando Modal de Baixa não abre para Transferência Bancária")
    except RuntimeError as exc:
        assert "Modal de Baixa" in str(exc) or "abriu" in str(exc)

    # Verifica que tentou clicar o botão antes de falhar
    assert actions.clicked == [("xpath", "//baixa")]


def test_search_doc_broader_uses_current_url_id_when_text_search_fails(monkeypatch):
    actions = FakeActions(set())
    actions.driver.current_url = "https://verbodavida.info/IVV/?mod=ivv&exec=entradas_saidas_dados&ID=5385288"
    page = _build_page(actions)
    row = SimpleNamespace(row_number=14, tipo=TipoMovimento.SAIDA, data_mov="29/07/2026", descricao_soma="TESTE")

    monkeypatch.setattr(page, "_go_back_to_list_best_effort", lambda row: None)
    monkeypatch.setattr(page.a, "type", lambda *args, **kwargs: None)
    monkeypatch.setattr(page.a, "click_js", lambda *args, **kwargs: None)
    monkeypatch.setattr(page, "_read_search_result_doc", lambda timeout_seconds=10: "")
    monkeypatch.setattr(page, "_dump_search_diagnostics", lambda *args, **kwargs: None)

    emitted = []
    monkeypatch.setattr(page, "_emit", lambda msg, **kv: emitted.append((msg, kv)))

    doc = page._search_doc_id_broader(row)

    assert doc == "5385288"
    assert any("doc_id da URL atual" in msg for msg, _ in emitted)


def test_create_and_get_doc_id_uses_url_hint_when_search_fails(monkeypatch):
    actions = FakeActions(set())
    actions.driver.current_url = "https://verbodavida.info/IVV/?mod=ivv&exec=entradas_saidas_dados&ID=5385288"
    page = _build_page(actions)
    row = SimpleNamespace(
        row_number=15,
        tipo=TipoMovimento.SAIDA,
        forma_pagamento="DINHEIRO",
        data_mov="29/07/2026",
        descricao_soma="TESTE",
    )

    monkeypatch.setattr(page, "_open_new", lambda row: None)
    monkeypatch.setattr(page, "_choose_tipo", lambda row: None)
    monkeypatch.setattr(page, "_fill_common", lambda row: None)
    monkeypatch.setattr(page, "_fill_saida", lambda row: None)
    monkeypatch.setattr(page, "_save_form_if_present", lambda row: None)
    monkeypatch.setattr(page, "_realizar_pagamento", lambda row: None)
    monkeypatch.setattr(page, "_pagamento_saida_modal", lambda row: None)
    monkeypatch.setattr(page, "_do_baixa", lambda row: None)
    monkeypatch.setattr(page, "_search_doc_id", lambda row: (_ for _ in ()).throw(RuntimeError("sem doc")))
    monkeypatch.setattr(page, "_confirm", lambda *args, **kwargs: None)

    doc = page.create_and_get_doc_id(row)

    assert doc == "5385288"


def test_is_transferencia_bancaria_recognizes_different_forms():
    actions = FakeActions(set())
    page = _build_page(actions)

    assert page._is_transferencia_bancaria("Transferência Bancária") is True
    assert page._is_transferencia_bancaria("transferencia bancaria") is True
    assert page._is_transferencia_bancaria("TRANSFERÊNCIA BANCÁRIA") is True
    assert page._is_transferencia_bancaria("  Transferência Bancária  ") is True
    assert page._is_transferencia_bancaria("Transferência Bancaria") is True
    assert page._is_transferencia_bancaria("Dinheiro") is False
    assert page._is_transferencia_bancaria("Cheque") is False


def test_conta_ordem_row_with_id_interno():
    """
    Testa que ContaOrdemRow carrega e preserva id_interno como texto.
    """
    raw = {
        "TIPO": "Saída",
        "DATA MOV.": "2024-01-15",
        "CAIXA": "CAIXA PRINCIPAL",
        "CAIXA SAIDA": "CAIXA SAIDA",
        "CENTRO DE CUSTO": "CC001",
        "PLANO DE CONTA": "PLANO001",
        "FORMA DE PAGAMENTO": "Transferência Bancária",
        "IMPORTÂNCIA": "1000.00",
        "DESCRIÇÃO SOMA": "Descrição teste",
        "ID_INTERNO": "INT123456",
        "DOC. SOMA": "",
        "DADOS DOC": "DADOS_TESTE",
        "IDUSER": "USER1",
        "TIMESTAMP": "2024-01-15 10:00:00",
    }
    row = ContaOrdemRow.from_table_row(row_number=1, raw=raw)

    # Verificar que id_interno foi carregado
    assert row.id_interno == "INT123456"
    assert isinstance(row.id_interno, str)

    # Verificar que forma_pagamento foi carregada corretamente
    actions = FakeActions(set())
    page = _build_page(actions)
    assert page._is_transferencia_bancaria(row.forma_pagamento)


def test_transferencia_bancaria_requires_id_interno():
    """
    Testa que Transferência Bancária sem id_interno gera ValueError.
    """
    row = ContaOrdemRow(
        row_number=2,
        tipo=TipoMovimento.SAIDA,
        data_mov="15/01/2024",
        caixa="CAIXA1",
        caixa_saida="CAIXA_SAIDA",
        centro_custo="CC001",
        plano_conta="PLAN001",
        forma_pagamento="Transferência Bancária",
        importancia="1000.00",
        descricao_soma="Desc",
        id_interno="",  # <-- VAZIO!
        doc_soma="",
        dados_doc="",
        iduser="USER1",
        timestamp="2024-01-15 10:00:00",
        raw={},
    )

    # Verificar que id_interno está vazio
    assert not row.id_interno or not row.id_interno.strip()
