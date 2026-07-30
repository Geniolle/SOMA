from __future__ import annotations

import json
from pathlib import Path
from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException

from soma_app.automation.pages.entradas_saidas_page import EntradasSaidasPage
from soma_app.config.fields import FORM_FIELD_REGISTRY, field_names


class FakeActions:
    def __init__(self, present):
        self.present = set(present)
        self.selected = []
        self.screenshots = []
        self.typed = []

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
