from __future__ import annotations

import json
from types import SimpleNamespace

import pytest
from selenium.webdriver.common.by import By

from soma_app.config.locators import (
    _coerce_locator,
    _coerce_locator_list,
    apply_locator_overrides,
    load_page_locator_config,
)


def test_coerce_locator_simple_string():
    raw = "//div[@id='test']"
    loc = _coerce_locator(raw, None)
    assert loc == (By.XPATH, "//div[@id='test']")


def test_coerce_locator_json_object():
    # css selector
    raw_css = {"by": "css", "value": ".class-name"}
    loc_css = _coerce_locator(raw_css, None)
    assert loc_css == (By.CSS_SELECTOR, ".class-name")

    # id
    raw_id = {"by": "id", "value": "element-id"}
    loc_id = _coerce_locator(raw_id, None)
    assert loc_id == (By.ID, "element-id")

    # name
    raw_name = {"by": "name", "value": "fieldName"}
    loc_name = _coerce_locator(raw_name, None)
    assert loc_name == (By.NAME, "fieldName")


def test_coerce_locator_list():
    raw_list = [
        "//first/xpath",
        {"by": "css", "value": ".second-selector"}
    ]
    locs = _coerce_locator_list(raw_list, [])
    assert locs == [
        (By.XPATH, "//first/xpath"),
        (By.CSS_SELECTOR, ".second-selector"),
    ]


class DummyPage:
    MANDATORY_LOCATOR = (By.XPATH, "")
    OPTIONAL_LOCATOR = (By.NAME, "default_name")
    MANDATORY_LIST = []
    FORM_READY_CANDIDATES = []


def test_apply_locator_overrides_raises_on_missing_mandatory(tmp_path):
    locators_path = tmp_path / "locators.json"
    # missing MANDATORY_LOCATOR
    locators_path.write_text(
        json.dumps(
            {
                "dummy_page": {
                    "MANDATORY_LIST": ["//list-item"]
                }
            }
        ),
        encoding="utf-8",
    )

    page = DummyPage()
    page.settings = SimpleNamespace(locators_path=locators_path)

    with pytest.raises(KeyError) as exc_info:
        apply_locator_overrides(page, "dummy_page")

    assert "MANDATORY_LOCATOR" in str(exc_info.value)


def test_apply_locator_overrides_raises_on_missing_mandatory_list(tmp_path):
    locators_path = tmp_path / "locators.json"
    # missing MANDATORY_LIST
    locators_path.write_text(
        json.dumps(
            {
                "dummy_page": {
                    "MANDATORY_LOCATOR": "//xpath"
                }
            }
        ),
        encoding="utf-8",
    )

    page = DummyPage()
    page.settings = SimpleNamespace(locators_path=locators_path)

    with pytest.raises(KeyError) as exc_info:
        apply_locator_overrides(page, "dummy_page")

    assert "MANDATORY_LIST" in str(exc_info.value)


def test_apply_locator_overrides_skips_derived_fields(tmp_path):
    locators_path = tmp_path / "locators.json"
    # missing FORM_READY_CANDIDATES (but it is derived, so shouldn't raise)
    locators_path.write_text(
        json.dumps(
            {
                "dummy_page": {
                    "MANDATORY_LOCATOR": "//xpath",
                    "MANDATORY_LIST": ["//list-item"]
                }
            }
        ),
        encoding="utf-8",
    )

    page = DummyPage()
    page.settings = SimpleNamespace(locators_path=locators_path)

    # Should not raise any error
    apply_locator_overrides(page, "dummy_page")
    assert page.MANDATORY_LOCATOR == (By.XPATH, "//xpath")
    assert page.MANDATORY_LIST == [(By.XPATH, "//list-item")]
    assert page.OPTIONAL_LOCATOR == (By.NAME, "default_name")  # default preserved


def test_load_real_locators_json_configuration():
    # Load config without settings to read default locators.json
    login_cfg = load_page_locator_config(None, "login")
    assert login_cfg is not None
    assert "SOMA_READY" in login_cfg
    assert "SOMA_BUTTON_CANDIDATES" in login_cfg
    assert "entradas_saidas" in login_cfg["SOMA_READY"]

    common_cfg = load_page_locator_config(None, "common")
    assert common_cfg is not None
    assert "SELECT2_SEARCH" in common_cfg
