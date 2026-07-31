from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from selenium.common.exceptions import TimeoutException

from soma_app.automation.pages.transferencias_page import TransferenciasPage
from soma_app.domain.models import TipoMovimento


class FakeTransferActions:
    def __init__(self, present, *, visible=None, enabled=None, texts=None):
        self.present = set(present)
        self.visible = set(visible if visible is not None else present)
        self.enabled = set(enabled if enabled is not None else present)
        self.texts = dict(texts or {})
        self.clicked = []
        self.js_clicked = []
        self.typed = []
        self.get_calls: list[str] = []
        self.driver = SimpleNamespace(
            current_url="https://test.local/",
            title="SOMA",
            get=self.get,
            execute_script=lambda *args, **kwargs: None,
        )

    def exists(self, locator, timeout_seconds=None):
        return locator in self.present

    def wait_any_present(self, locators, timeout_seconds=None):
        for loc in locators:
            if loc in self.present:
                return loc
        raise TimeoutException("missing")

    def wait_visible(self, locator, timeout_seconds=None):
        if locator in self.present and locator in self.visible:
            return SimpleNamespace(text=self.texts.get(locator, ""), is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("missing")

    def wait_present(self, locator, timeout_seconds=None):
        if locator in self.present:
            return SimpleNamespace(text=self.texts.get(locator, ""), is_displayed=lambda: True, is_enabled=lambda: True)
        raise TimeoutException("missing")

    def wait_dom_ready(self, timeout_seconds=30):
        return None

    def get(self, url):
        self.get_calls.append(url)

    def click(self, locator):
        self.clicked.append(locator)

    def click_js(self, locator):
        self.js_clicked.append(locator)

    def type(self, locator, value, clear=True):
        self.typed.append((locator, value, clear))

    def screenshot(self, name):
        return Path(f"{name}.png")


def _build_page(actions, settings=None):
    settings = settings or SimpleNamespace(site_home_url="https://verbodavida.info/IVV/", timeout_seconds=5)
    page = TransferenciasPage(actions, settings)
    return page


def test_transferencias_run_follows_open_fill_and_back():
    menu = ("xpath", "//menu")
    nova = ("xpath", "//nova")
    caixa_saida = ("xpath", "//caixa-saida")
    caixa_entrada = ("xpath", "//caixa-entrada")
    valor = ("xpath", "//valor")
    data = ("xpath", "//data")
    descricao = ("xpath", "//descricao")
    salvar = ("xpath", "//salvar")
    voltar = ("xpath", "//voltar")
    search = ("xpath", "//search")

    actions = FakeTransferActions({menu, nova, caixa_saida, caixa_entrada, valor, data, descricao, salvar, voltar, search})
    page = _build_page(actions)
    page.MENU_CAIXAS_BANCOS_CANDIDATES = [menu]
    page.BTN_NOVA_TRANSFERENCIA_CANDIDATES = [nova]
    page.CAIXA_SAIDA = caixa_saida
    page.CAIXA_ENTRADA = caixa_entrada
    page.VALOR = valor
    page.DATA = data
    page.DESCRICAO = descricao
    page.BTN_SALVAR = salvar
    page.BTN_VOLTAR = voltar
    page.SELECT2_SEARCH = search

    row = SimpleNamespace(
        row_number=3,
        tipo=TipoMovimento.TRANSFERENCIA,
        caixa_saida="CAIXA A",
        caixa="CAIXA B",
        importancia="10,00",
        data_mov="29/07/2026",
        descricao_soma="transferencia teste",
    )

    def select2(opener, value, *, row, field):
        actions.clicked.append(opener)

    page._select2_choose_verified = select2

    result = page.run(row)

    assert result == "Transferido"
    assert actions.get_calls == ["https://verbodavida.info/IVV/"]
    assert actions.js_clicked == [menu, nova, salvar, voltar]
    assert actions.typed


def test_transferencias_select2_variants_strip_accents():
    variants = TransferenciasPage._select2_variants("CAIXA DIÁRIO")
    assert "CAIXA DIÁRIO" in variants
    assert "CAIXA DIARIO" in variants


def test_transferencias_select2_match_index_prefers_normalized_match():
    options = ["VERBO SHOP", "VERBO CAFÉ", "D. CRIANÇAS", "CAIXA ECONÔMICA", "CAIXA DIÁRIO"]

    idx = TransferenciasPage._select2_match_index(options, "CAIXA DIARIO")

    assert idx == 4
