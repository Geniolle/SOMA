import unittest

from soma_app.domain.models import ContaOrdemRow, LinhaStatus, TipoMovimento
from soma_app.domain.rules import should_process, validate_row


class TestRules(unittest.TestCase):
    def _base_row(self, tipo: str, doc: str = "") -> ContaOrdemRow:
        raw = {
            "TIPO": tipo,
            "DATA MOV.": "2026-02-27",
            "CAIXA": "CAIXA_A",
            "CAIXA SAIDA": "CAIXA_B",
            "CENTRO DE CUSTO": "CC1",
            "PLANO DE CONTA": "PC1",
            "FORMA DE PAGAMENTO": "DINHEIRO",
            "IMPORTÂNCIA": "1",
            "DESCRIÇÃO SOMA": "descr",
            "DOC. SOMA": doc,
        }
        return ContaOrdemRow.from_table_row(2, raw)

    def test_should_process(self):
        self.assertTrue(should_process(self._base_row("Entrada", "")))
        self.assertTrue(should_process(self._base_row("Entrada", "Em processamento")))
        self.assertFalse(should_process(self._base_row("Entrada", "12345")))
        self.assertFalse(should_process(self._base_row("Entrada", "Em erro")))
        self.assertTrue(should_process(self._base_row("Entrada", "Em erro"), allow_retry_error=True))

    def test_validate_saida_requires_caixa_saida(self):
        raw = {
            "TIPO": "Saída",
            "DATA MOV.": "2026-02-27",
            "CAIXA": "CAIXA_A",
            "CAIXA SAIDA": "",
            "CENTRO DE CUSTO": "CC1",
            "PLANO DE CONTA": "PC1",
            "FORMA DE PAGAMENTO": "DINHEIRO",
            "IMPORTÂNCIA": "1",
            "DESCRIÇÃO SOMA": "descr",
            "DOC. SOMA": "",
        }
        row = ContaOrdemRow.from_table_row(5, raw)
        errs = validate_row(row)
        self.assertTrue(any(e.field == "CAIXA SAIDA" for e in errs))

    def test_validate_transfer_requires_caixa_saida(self):
        raw = {
            "TIPO": "Transferência",
            "DATA MOV.": "2026-02-27",
            "CAIXA": "CAIXA_A",
            "CAIXA SAIDA": "",
            "CENTRO DE CUSTO": "CC1",
            "PLANO DE CONTA": "PC1",
            "FORMA DE PAGAMENTO": "DINHEIRO",
            "IMPORTÂNCIA": "1",
            "DESCRIÇÃO SOMA": "descr",
            "DOC. SOMA": "",
        }
        row = ContaOrdemRow.from_table_row(6, raw)
        errs = validate_row(row)
        self.assertTrue(any(e.field == "CAIXA SAIDA" for e in errs))


if __name__ == "__main__":
    unittest.main()