from __future__ import annotations

from pathlib import Path

from soma_app.automation.dom_inventory import DomInventory
from soma_app.tools._dom_inventory_patch import apply_dom_inventory_javascript_fix


class ScriptCapturingDriver:
    def __init__(self) -> None:
        self.script = ""

    def execute_script(self, script: str):
        self.script = script
        return []


def test_dom_inventory_patch_uses_compilable_quote_free_xpath_builder():
    apply_dom_inventory_javascript_fix()
    driver = ScriptCapturingDriver()
    inventory = DomInventory()

    result = inventory._collect_context_payload(driver)

    assert result == {"elements": []}
    assert "escapeXPathLiteral" not in driver.script
    assert 'join(", "\'", ")' not in driver.script
    assert "segments.join('/')" in driver.script
    assert "__INTERACTIVE_SELECTOR__" not in driver.script


def test_tools_package_does_not_import_site_mapper_early():
    package_init = Path(__file__).resolve().parents[2] / "src" / "soma_app" / "tools" / "__init__.py"
    source = package_init.read_text(encoding="utf-8")

    assert "from .site_mapper import" not in source
