from __future__ import annotations

import inspect
import json
import logging
from collections import Counter
from pathlib import Path
from types import SimpleNamespace

import pytest
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By

import soma_app.automation.pages.entradas_saidas_page as entradas_saidas_page_module
from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.domain.models import TipoMovimento


class FakeActions:
    def __init__(self, present):
        self.present = set(present)
        self.selected = []
        self.screenshots = []
        self.typed = []
        self.debug_context = None
        self.driver = SimpleNamespace(
            current_url="http://example.invalid",
            execute_script=lambda script, element: None,
        )

    def exists(self, locator, timeout_seconds=None):
        return locator in self.present

    def wait_any_present(self, locators, timeout_seconds=None):
        for loc in locators:
            if loc in self.present:
                return loc
        raise TimeoutException("missing")

    def select2_choose(self, locator, value):
        self.selected.append((locator, value))

    def type(self, locator, value, clear=True):
        self.typed.append((locator, value, clear))

    def click_js(self, locator):
        return None

    def click_any_visible(self, locators, timeout_seconds=None):
        for loc in locators:
            if loc in self.present:
                self.selected.append((loc, "__click_any_visible__"))
                return None
        return None

    def wait_visible(self, locator, timeout_seconds=None):
        raise TimeoutException("missing")

    def wait_invisible(self, locator, timeout_seconds=None):
        return None

    def wait_any_visible_element(self, locators, timeout_seconds=None, log_timeout=True):
        return self.wait_any_present(locators, timeout_seconds=timeout_seconds)

    def set_debug_context(self, context):
        self.debug_context = context

    def screenshot(self, name):
        self.screenshots.append(name)
        return Path(f"{name}.png")

    def dump_page_source(self, name):
        return Path(f"{name}.html")


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
    entradas_saidas_cfg = {}
    for name, val in vars(EntradasSaidasPage).items():
        if name.isupper() and name not in {"RADIO_ANY_CANDIDATES", "FORM_READY_CANDIDATES", "ANY_VALUE_CANDIDATES"}:
            if _is_locator(val):
                entradas_saidas_cfg[name] = f"//dummy-{name.lower()}"
            elif _is_locator_list(val):
                entradas_saidas_cfg[name] = [f"//dummy-{name.lower()}"]

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


def test_fill_caixa_entrada_ultima_does_not_raise_when_selection_changes(monkeypatch):
    from types import SimpleNamespace

    actions = FakeActions({("xpath", "//caixa-container")})
    page = _build_page(actions)
    page.CAIXA_ENTRADA_CONTAINER = ("xpath", "//caixa-container")
    row = SimpleNamespace(row_number=12, tipo=SimpleNamespace(value="Entrada"), caixa="CAIXA DIÁRIO")

    monkeypatch.setattr(page, "_select_with_sleep_validation", lambda *args, **kwargs: "CAIXA ECONÔMICA MONTEPIO GERAL [CONTA CORRENTE]")

    page._fill_caixa_entrada_ultima(row)


def test_realizar_pagamento_is_best_effort_when_button_is_missing():
    from types import SimpleNamespace

    class TimeoutClickActions(FakeActions):
        def click_js(self, locator):
            raise TimeoutException("missing")

    actions = TimeoutClickActions(set())
    page = _build_page(actions)
    row = SimpleNamespace(row_number=42, tipo=SimpleNamespace(value="Entrada"))

    page._realizar_pagamento(row)


def test_search_doc_id_attempt_uses_fallback_when_lookup_returns_none(monkeypatch):
    from types import SimpleNamespace

    actions = FakeActions(set())
    page = _build_page(actions)
    row = SimpleNamespace(row_number=99, tipo=SimpleNamespace(value="Entrada"))
    monkeypatch.setattr(page, "_search_doc_lookup_attempt", lambda r: None)

    doc = page._search_doc_id_attempt(row)

    assert doc == "SEM_DOC_99"


def test_search_existing_doc_reuses_lookup_attempt(monkeypatch):
    from types import SimpleNamespace

    actions = FakeActions(set())
    page = _build_page(actions)
    row = SimpleNamespace(row_number=98, tipo=SimpleNamespace(value="Entrada"))
    monkeypatch.setattr(page, "_search_doc_lookup_attempt", lambda r: "DOC-EXIST")

    doc = page.search_existing_doc(row)

    assert doc == "DOC-EXIST"


def test_save_form_if_present_uses_direct_click_without_long_wait(monkeypatch):
    from types import SimpleNamespace

    class FakeButton:
        def __init__(self):
            self.clicked = False

    class FakeDriver:
        def __init__(self):
            self.button = FakeButton()
            self.executed = []

        def find_element(self, by, selector):
            return self.button

        def execute_script(self, script, element):
            self.executed.append(script)
            element.clicked = True

    class SaveActions(FakeActions):
        def __init__(self):
            super().__init__({("xpath", "//salvar")})
            self.driver = FakeDriver()

    actions = SaveActions()
    page = _build_page(actions)
    page.BTN_SALVAR_FORM = ("xpath", "//salvar")
    row = SimpleNamespace(row_number=7, tipo=SimpleNamespace(value="Saída"))

    page._save_form_if_present(row)

    assert actions.driver.button.clicked is True
    assert actions.driver.executed == ["arguments[0].click();"]


def test_saida_payment_strict_methods_do_not_use_candidates_or_best_effort():
    src_pagamento = inspect.getsource(EntradasSaidasPage._pagamento_saida_modal_strict)
    src_baixa = inspect.getsource(EntradasSaidasPage._inserir_baixa_saida)

    assert "_caixa_pagamento_modal_locator" in src_pagamento
    assert "_select_fixed_visible_text" in src_pagamento
    assert "click_any_visible" in src_baixa
    assert "wait_any_visible_element" in src_baixa


def test_saida_payment_fixed_xpaths_are_absolute():
    page = _build_page(FakeActions(set()))

    assert page.PLANO_CONTA == ("xpath", "/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span/span[1]")
    assert page.BTN_INSERIR_PAGAMENTO_SAIDA == ("xpath", "//*[@data-target='#inserir' and contains(normalize-space(.), 'Inserir Pagamento')]")
    assert page.DATA_PAGAMENTO_MODAL == ("xpath", "//*[@id='inserir']//input[@name='data_pagamento']")
    assert page.FORMA_PAGAMENTO_MODAL == ("xpath", "//*[@id='inserir']//select[@name='forma_pagamento']")
    assert page.CAIXA_PAGAMENTO_MODAL == ("xpath", "//*[@id='inserir']//select[@name='id_caixa_origem']")
    assert page.OK_ALERT == ("css selector", ".swal2-confirm")
    assert page.NUM_DOCUMENTO_MODAL == ("xpath", "//*[@id='inserir']//input[@name='num_documento']")
    assert page.BTN_SALVAR_PAGAMENTO_MODAL == ("xpath", "//*[@id='inserir']//button[@id='botao_pagamento' or contains(normalize-space(.), 'Salvar')]")
    assert page.BTN_INSERIR_BAIXA == ("xpath", "//*[@data-target='#inserirBaixa' or contains(@data-target, 'inserirBaixa')]")
    assert page.DATA_BAIXA == ("xpath", "//*[@id='inserirBaixa']//input[@name='data_baixa']")
    assert page.BTN_SALVAR_BAIXA == ("xpath", "//*[@id='inserirBaixa']//button[@id='botao_baixa' or contains(normalize-space(.), 'Salvar Baixa')]")


def test_locators_json_separates_payment_and_baixa_without_duplicate_relevant_keys():
    path = Path(__file__).resolve().parents[1] / "src" / "soma_app" / "config" / "locators.json"
    duplicate_keys = []

    def hook(pairs):
        counts = Counter(key for key, _ in pairs)
        duplicate_keys.extend(key for key, count in counts.items() if count > 1)
        return dict(pairs)

    data = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=hook)
    es = data["entradas_saidas"]

    assert es["PLANO_CONTA"] == "/html/body/div[2]/div/div[3]/div/div/form/div[7]/div/span/span[1]/span/span[1]"
    assert es["BTN_INSERIR_BAIXA"] == "//*[@data-target='#inserirBaixa' or contains(@data-target, 'inserirBaixa')]"
    assert es["DATA_BAIXA"] == "//*[@id='inserirBaixa']//input[@name='data_baixa']"
    assert es["BTN_SALVAR_BAIXA"] == "//*[@id='inserirBaixa']//button[@id='botao_baixa' or contains(normalize-space(.), 'Salvar Baixa')]"
    assert "FORMA_PAGAMENTO_MODAL" not in duplicate_keys
    assert "PLANO_CONTA" not in duplicate_keys
    assert "BTN_INSERIR_BAIXA" not in duplicate_keys
    assert "DATA_BAIXA" not in duplicate_keys
    assert "BTN_SALVAR_BAIXA" not in duplicate_keys


def test_numero_documento_transferencia_usa_id_interno_e_rejeita_fallback():
    page = _build_page(FakeActions(set()))

    row = SimpleNamespace(row_number=8, forma_pagamento="TRANSFERÊNCIA BANCÁRIA", id_interno="INT-99", descricao_soma="DESC-IGNORADA")
    assert page._numero_documento_para_pagamento_saida(row) == "INT-99"

    row_vazio = SimpleNamespace(row_number=8, forma_pagamento="TRANSFERÊNCIA BANCÁRIA", id_interno="", descricao_soma="DESC-IGNORADA")
    with pytest.raises(ValueError) as exc_info:
        page._numero_documento_para_pagamento_saida(row_vazio)
    assert "ID_INTERNO" in str(exc_info.value)


def test_select_fixed_visible_text_retries_same_xpath_when_select_goes_stale(monkeypatch):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(row_number=21, tipo=SimpleNamespace(value="Saída"))
    elements = [object(), object(), object()]
    calls = []

    class FakeSelect:
        def __init__(self, element):
            self.element = element

        def select_by_visible_text(self, value):
            calls.append(("select", self.element, value))
            if self.element in elements[:2]:
                raise StaleElementReferenceException("stale select")

        @property
        def first_selected_option(self):
            class Opt:
                text = "CAIXA CENTRAL"

            return Opt()

    def _wait_fixed_visible(locator, **kwargs):
        calls.append(("wait", locator))
        return elements[len([c for c in calls if c[0] == "wait"]) - 1]

    monkeypatch.setattr(entradas_saidas_page_module, "Select", FakeSelect)
    monkeypatch.setattr(page, "_wait_fixed_visible", _wait_fixed_visible)

    result = page._select_fixed_visible_text(
        page.CAIXA_PAGAMENTO_MODAL,
        "CAIXA CENTRAL",
        row=row,
        stage="entradas_saidas.baixa.caixa",
        element_name="CAIXA_PAGAMENTO_MODAL",
        timeout_seconds=1,
        stale_retries=3,
        stale_label="CAIXA",
    )

    assert result == "CAIXA CENTRAL"
    assert [c for c in calls if c[0] == "wait"] == [
        ("wait", page.CAIXA_PAGAMENTO_MODAL),
        ("wait", page.CAIXA_PAGAMENTO_MODAL),
        ("wait", page.CAIXA_PAGAMENTO_MODAL),
    ]


def test_select_fixed_visible_text_retries_when_validation_goes_stale(monkeypatch):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(row_number=22, tipo=SimpleNamespace(value="Saída"))
    elements = [object(), object(), object()]
    calls = []

    class FakeSelect:
        def __init__(self, element):
            self.element = element

        def select_by_visible_text(self, value):
            calls.append(("select", self.element, value))

        @property
        def first_selected_option(self):
            if self.element in elements[:2]:
                raise StaleElementReferenceException("stale validation")

            class Opt:
                text = "CAIXA CENTRAL"

            return Opt()

    def _wait_fixed_visible(locator, **kwargs):
        calls.append(("wait", locator))
        return elements[len([c for c in calls if c[0] == "wait"]) - 1]

    monkeypatch.setattr(entradas_saidas_page_module, "Select", FakeSelect)
    monkeypatch.setattr(page, "_wait_fixed_visible", _wait_fixed_visible)

    result = page._select_fixed_visible_text(
        page.CAIXA_PAGAMENTO_MODAL,
        "CAIXA CENTRAL",
        row=row,
        stage="entradas_saidas.baixa.caixa",
        element_name="CAIXA_PAGAMENTO_MODAL",
        timeout_seconds=1,
        stale_retries=3,
        stale_label="CAIXA",
    )

    assert result == "CAIXA CENTRAL"
    assert [c for c in calls if c[0] == "wait"] == [
        ("wait", page.CAIXA_PAGAMENTO_MODAL),
        ("wait", page.CAIXA_PAGAMENTO_MODAL),
        ("wait", page.CAIXA_PAGAMENTO_MODAL),
    ]


def test_select_fixed_visible_text_exhausts_retries_with_explicit_stale_error(monkeypatch, caplog):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(row_number=23, tipo=SimpleNamespace(value="Saída"))
    elements = [object(), object(), object()]
    calls = []

    class FakeSelect:
        def __init__(self, element):
            self.element = element

        def select_by_visible_text(self, value):
            raise StaleElementReferenceException("stale select")

        @property
        def first_selected_option(self):
            class Opt:
                text = "CAIXA CENTRAL"

            return Opt()

    def _wait_fixed_visible(locator, **kwargs):
        calls.append(("wait", locator))
        return elements[len([c for c in calls if c[0] == "wait"]) - 1]

    monkeypatch.setattr(entradas_saidas_page_module, "Select", FakeSelect)
    monkeypatch.setattr(page, "_wait_fixed_visible", _wait_fixed_visible)

    with caplog.at_level(logging.ERROR):
        with pytest.raises(TimeoutException) as exc_info:
            page._select_fixed_visible_text(
                page.CAIXA_PAGAMENTO_MODAL,
                "CAIXA CENTRAL",
                row=row,
                stage="entradas_saidas.baixa.caixa",
                element_name="CAIXA_PAGAMENTO_MODAL",
                timeout_seconds=1,
                stale_retries=3,
                stale_label="CAIXA",
            )

    assert "[SAIDA][XPATH_ERROR]" in caplog.text
    assert "causa=StaleElementReferenceException" in caplog.text
    assert "elemento=CAIXA_PAGAMENTO_MODAL" in caplog.text
    assert page.CAIXA_PAGAMENTO_MODAL[1] in caplog.text
    assert len([c for c in calls if c[0] == "wait"]) == 3
    assert "screenshot=" in str(exc_info.value)


def test_pagamento_saida_modal_strict_uses_fixed_xpaths_and_row_values(monkeypatch):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(
        row_number=12,
        tipo=SimpleNamespace(value="Saída"),
        data_mov="12/08/2026",
        forma_pagamento="DINHEIRO",
        caixa="CAIXA CENTRAL",
        id_interno="INT-12",
    )

    calls = []
    ok_el = object()

    monkeypatch.setattr(page, "_dismiss_overlays", lambda: None)
    monkeypatch.setattr(page, "_click_fixed_visible", lambda locator, **kwargs: calls.append(("click", locator, kwargs["element_name"], kwargs["stage"])))
    monkeypatch.setattr(page, "_type_fixed_visible", lambda locator, value, **kwargs: calls.append(("type", locator, value, kwargs["element_name"], kwargs["stage"])) or value)
    monkeypatch.setattr(page, "_select_fixed_visible_text", lambda locator, value, **kwargs: calls.append(("select", locator, value, kwargs["element_name"], kwargs["stage"])) or value)
    monkeypatch.setattr(page, "_wait_fixed_displayed", lambda locator, **kwargs: ok_el)
    monkeypatch.setattr(page.a.driver, "execute_script", lambda script, element: calls.append(("execute_script", element)))
    monkeypatch.setattr(page.a, "wait_invisible", lambda locator, **kwargs: calls.append(("wait_invisible", locator)))

    page._pagamento_saida_modal_strict(row)

    assert calls == [
        ("click", page.BTN_INSERIR_PAGAMENTO_SAIDA, "BTN_INSERIR_PAGAMENTO_SAIDA", "entradas_saidas.pagamento_saida_modal.open"),
        ("type", page.DATA_PAGAMENTO_MODAL, "12/08/2026", "DATA_PAGAMENTO_MODAL", "entradas_saidas.pagamento_saida_modal.data_pagamento"),
        ("select", page.FORMA_PAGAMENTO_MODAL, "DINHEIRO", "FORMA_PAGAMENTO_MODAL", "entradas_saidas.pagamento_saida_modal.forma_pagamento"),
        ("select", page.CAIXA_PAGAMENTO_MODAL, "CAIXA CENTRAL", "CAIXA_PAGAMENTO_MODAL", "entradas_saidas.pagamento_saida_modal.caixa"),
        ("click", page.BTN_SALVAR_PAGAMENTO_MODAL, "BTN_SALVAR_PAGAMENTO_MODAL", "entradas_saidas.pagamento_saida_modal.salvar"),
        ("execute_script", ok_el),
        ("wait_invisible", page.OK_ALERT),
    ]
    assert all(entry[0] != "click" or entry[1] is not page.OK_ALERT for entry in calls)


def test_pagamento_saida_modal_strict_transferencia_usa_id_interno(monkeypatch):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(
        row_number=13,
        tipo=SimpleNamespace(value="Saída"),
        data_mov="12/08/2026",
        forma_pagamento="TRANSFERÊNCIA BANCÁRIA",
        caixa="CAIXA CENTRAL",
        id_interno="ID-ABC",
        descricao_soma="DESC-IGNORADA",
    )

    typed = []
    calls = []
    ok_el = object()

    monkeypatch.setattr(page, "_dismiss_overlays", lambda: None)
    monkeypatch.setattr(page, "_click_fixed_visible", lambda locator, **kwargs: None)
    monkeypatch.setattr(page, "_select_fixed_visible_text", lambda locator, value, **kwargs: value)
    monkeypatch.setattr(page, "_wait_fixed_displayed", lambda locator, **kwargs: ok_el)
    monkeypatch.setattr(page.a.driver, "execute_script", lambda script, element: calls.append(("execute_script", element)))
    monkeypatch.setattr(page.a, "wait_invisible", lambda locator, **kwargs: calls.append(("wait_invisible", locator)))

    def _capture_type(locator, value, **kwargs):
        typed.append((locator, value, kwargs["element_name"]))
        return value

    monkeypatch.setattr(page, "_type_fixed_visible", _capture_type)

    page._pagamento_saida_modal_strict(row)

    assert (page.NUM_DOCUMENTO_MODAL, "ID-ABC", "NUM_DOCUMENTO_MODAL") in typed
    assert all(value != "DESC-IGNORADA" for _, value, _ in typed)
    assert ("execute_script", ok_el) in calls
    assert ("wait_invisible", page.OK_ALERT) in calls


def test_inserir_baixa_saida_strict_uses_fixed_xpaths_and_does_not_dismiss_overlays(monkeypatch):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(
        row_number=15,
        tipo=SimpleNamespace(value="Saída"),
        data_mov="12/08/2026",
        forma_pagamento="DINHEIRO",
        caixa="CAIXA CENTRAL",
        id_interno="ID-15",
    )

    calls = []

    def _fail_if_dismissed():
        raise AssertionError("_dismiss_overlays nao deve ser chamado neste trecho")

    fake_el = SimpleNamespace(
        clear=lambda: None,
        send_keys=lambda *_args, **_kwargs: None,
        get_attribute=lambda name: "12/08/2026" if name == "value" else None,
    )

    monkeypatch.setattr(page, "_dismiss_overlays", _fail_if_dismissed)
    monkeypatch.setattr(page, "_click_fixed_visible", lambda *args, **kwargs: None)
    monkeypatch.setattr(page, "_type_fixed_visible", lambda locator, value, **kwargs: calls.append(("type", locator, value, kwargs["element_name"], kwargs["stage"])) or value)
    monkeypatch.setattr(page, "_select_fixed_visible_text", lambda locator, value, **kwargs: calls.append(("select", locator, value, kwargs["element_name"], kwargs["stage"])) or value)
    monkeypatch.setattr(page.a, "wait_any_visible_element", lambda locators, **kwargs: fake_el)
    monkeypatch.setattr(page.a, "wait_visible", lambda locator, **kwargs: fake_el)
    monkeypatch.setattr(page.a, "wait_invisible", lambda locator, **kwargs: calls.append(("wait_invisible", locator)))

    page._inserir_baixa_saida(row)

    assert ("wait_invisible", (By.ID, "inserirBaixa")) in calls
    assert not any(entry[0] == "dismiss" for entry in calls)


def test_inserir_baixa_saida_strict_timeout_on_save_raises_xpath_error(monkeypatch, caplog):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(
        row_number=16,
        tipo=SimpleNamespace(value="Saída"),
        data_mov="12/08/2026",
        forma_pagamento="DINHEIRO",
        caixa="CAIXA CENTRAL",
        id_interno="ID-16",
    )

    monkeypatch.setattr(page, "_dismiss_overlays", lambda: None)
    monkeypatch.setattr(page, "_click_fixed_visible", lambda *args, **kwargs: None)
    monkeypatch.setattr(page, "_type_fixed_visible", lambda locator, value, **kwargs: value)
    monkeypatch.setattr(page, "_select_fixed_visible_text", lambda locator, value, **kwargs: value)
    fake_el = SimpleNamespace(
        clear=lambda: None,
        send_keys=lambda *_args, **_kwargs: None,
        get_attribute=lambda name: "12/08/2026" if name == "value" else None,
    )
    monkeypatch.setattr(page.a, "wait_any_visible_element", lambda locators, **kwargs: fake_el)
    monkeypatch.setattr(page.a, "wait_visible", lambda locator, **kwargs: fake_el)
    monkeypatch.setattr(page.a, "wait_invisible", lambda locator, **kwargs: (_ for _ in ()).throw(TimeoutException("missing")))

    with caplog.at_level(logging.ERROR):
        with pytest.raises(TimeoutException) as exc_info:
            page._inserir_baixa_saida(row)

    assert "[SAIDA][XPATH_ERROR]" in caplog.text
    assert "linha=16" in caplog.text
    assert "elemento=BTN_SALVAR_BAIXA" in caplog.text
    assert page.BTN_SALVAR_BAIXA_CANDIDATES[0][1] in caplog.text
    assert "etapa=entradas_saidas.baixa.salvar" in caplog.text
    assert "url=http://example.invalid" in caplog.text
    assert "causa=BAIXA_NAO_CONCLUIDA" in caplog.text
    assert "screenshot=" in str(exc_info.value)


def test_saida_payment_debug_checkpoint_for_save_and_confirmacao_is_present():
    src_pagamento = inspect.getsource(EntradasSaidasPage._pagamento_saida_modal_strict)
    src_baixa = inspect.getsource(EntradasSaidasPage._inserir_baixa_saida)

    assert "SAIDA.PAGAMENTO.SALVAR" in src_pagamento
    assert "SAIDA.PAGAMENTO.CONFIRMACAO" in src_pagamento
    assert "/html/body/div[5]/div/button[1]" in src_pagamento
    assert "SAIDA.BAIXA.INSERIR" in src_baixa
    assert "SAIDA.BAIXA.DATA" in src_baixa
    assert "SAIDA.BAIXA.SALVAR" in src_baixa


def test_fixed_xpath_failure_logs_context(caplog, monkeypatch):
    page = _build_page(FakeActions(set()))
    row = SimpleNamespace(row_number=123, tipo=SimpleNamespace(value="Saída"))

    def _missing(*args, **kwargs):
        raise TimeoutException("missing")

    monkeypatch.setattr(page.a, "wait_visible", _missing)

    with caplog.at_level(logging.ERROR):
        with pytest.raises(TimeoutException):
            page._wait_fixed_visible(
                page.DATA_PAGAMENTO_MODAL,
                row=row,
                stage="entradas_saidas.pagamento_saida_modal.data_pagamento",
                element_name="DATA_PAGAMENTO_MODAL",
                timeout_seconds=1,
            )

    assert "[SAIDA][XPATH_ERROR]" in caplog.text
    assert "linha=123" in caplog.text
    assert "elemento=DATA_PAGAMENTO_MODAL" in caplog.text
    assert page.DATA_PAGAMENTO_MODAL[1] in caplog.text


def test_create_and_get_doc_id_saida_executes_payment_then_baixa_before_search(monkeypatch):
    page = _build_page(FakeActions(set()))
    page.a.present.add(page.BTN_REALIZAR_PAGAMENTO)
    row = SimpleNamespace(
        row_number=44,
        tipo=TipoMovimento.SAIDA,
        forma_pagamento="DINHEIRO",
        data_mov="12/08/2026",
        caixa="CAIXA CENTRAL",
    )

    calls = []
    monkeypatch.setattr(page, "_open_new", lambda r: calls.append("open"))
    monkeypatch.setattr(page, "_choose_tipo", lambda r: calls.append("tipo"))
    monkeypatch.setattr(page, "_fill_common", lambda r: calls.append("common"))
    monkeypatch.setattr(page, "_fill_saida", lambda r: calls.append("saida"))
    monkeypatch.setattr(page, "_save_form_if_present", lambda r: calls.append("save_form"))
    monkeypatch.setattr(page, "_realizar_pagamento", lambda r: calls.append("realizar"))
    monkeypatch.setattr(page, "_inserir_pagamento_saida", lambda r: calls.append("pagamento_strict"))
    monkeypatch.setattr(page, "_inserir_baixa_saida", lambda r: calls.append("baixa_strict"))
    monkeypatch.setattr(page, "_search_doc_id", lambda r: calls.append("search") or "DOC-OK")

    doc = page.create_and_get_doc_id(row)

    assert doc == "DOC-OK"
    assert calls == ["open", "tipo", "common", "saida", "save_form", "realizar", "pagamento_strict", "baixa_strict", "search"]


def test_create_and_get_doc_id_saida_does_not_search_doc_when_baixa_fails(monkeypatch):
    page = _build_page(FakeActions(set()))
    page.a.present.add(page.BTN_REALIZAR_PAGAMENTO)
    row = SimpleNamespace(
        row_number=45,
        tipo=TipoMovimento.SAIDA,
        forma_pagamento="DINHEIRO",
        data_mov="12/08/2026",
        caixa="CAIXA CENTRAL",
    )

    calls = []
    monkeypatch.setattr(page, "_open_new", lambda r: calls.append("open"))
    monkeypatch.setattr(page, "_choose_tipo", lambda r: calls.append("tipo"))
    monkeypatch.setattr(page, "_fill_common", lambda r: calls.append("common"))
    monkeypatch.setattr(page, "_fill_saida", lambda r: calls.append("saida"))
    monkeypatch.setattr(page, "_save_form_if_present", lambda r: calls.append("save_form"))
    monkeypatch.setattr(page, "_realizar_pagamento", lambda r: calls.append("realizar"))
    monkeypatch.setattr(page, "_inserir_pagamento_saida", lambda r: calls.append("pagamento_strict"))
    def _fail_baixa(r):
        calls.append("baixa_strict")
        raise RuntimeError("baixa falhou")
    monkeypatch.setattr(page, "_inserir_baixa_saida", _fail_baixa)
    monkeypatch.setattr(page, "_search_doc_id", lambda r: calls.append("search") or "DOC-OK")

    with pytest.raises(RuntimeError, match="baixa falhou"):
        page.create_and_get_doc_id(row)

    assert calls == ["open", "tipo", "common", "saida", "save_form", "realizar", "pagamento_strict", "baixa_strict"]


def test_resolve_caixa_pagamento_modal_select_uses_origin_first():
    actions = FakeActions(set())
    origin_locator = ("xpath", "//*[@id='inserir']//select[@name='id_caixa_origem']")
    actions.present.add(origin_locator)
    seen = []
    actions.driver.find_element = lambda *args: seen.append(args) or object()
    page = _build_page(actions)

    result = page._resolve_caixa_pagamento_modal_select()

    assert result is not None
    assert seen[0] == origin_locator
