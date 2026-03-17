import unittest

from soma_app.domain.models import (
    ContaOrdemRow,
    LinhaStatus,
    TipoMovimento,
    WorkflowKind,
    status_from_doc_soma,
)


class TestParsing(unittest.TestCase):
    def test_tipo_normalization(self):
        self.assertEqual(TipoMovimento.from_sheet_value("Saída"), TipoMovimento.SAIDA)
        self.assertEqual(TipoMovimento.from_sheet_value("Saida"), TipoMovimento.SAIDA)
        self.assertEqual(TipoMovimento.from_sheet_value("transferencia"), TipoMovimento.TRANSFERENCIA)
        self.assertEqual(TipoMovimento.from_sheet_value("Entrada"), TipoMovimento.ENTRADA)

    def test_status_from_doc(self):
        self.assertEqual(status_from_doc_soma(""), LinhaStatus.PENDENTE)
        self.assertEqual(status_from_doc_soma("Em processamento"), LinhaStatus.EM_PROCESSAMENTO)
        self.assertEqual(status_from_doc_soma("Em erro"), LinhaStatus.EM_ERRO)
        self.assertEqual(status_from_doc_soma("12345"), LinhaStatus.CONCLUIDA)

    def test_workflow_selection(self):
        r = ContaOrdemRow.from_table_row(
            10,
            {
                "TIPO": "Transferência",
                "DATA MOV.": "2026-02-27",
                "CAIXA": "A",
                "CAIXA SAIDA": "B",
                "CENTRO DE CUSTO": "X",
                "PLANO DE CONTA": "Y",
                "FORMA DE PAGAMENTO": "DINHEIRO",
                "IMPORTÂNCIA": "1",
                "DESCRIÇÃO SOMA": "teste",
                "DOC. SOMA": "",
            },
        )
        self.assertEqual(r.tipo, TipoMovimento.TRANSFERENCIA)
        self.assertEqual(r.workflow, WorkflowKind.TRANSFERENCIAS)


if __name__ == "__main__":
    unittest.main()