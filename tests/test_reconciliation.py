"""Testes para soma_app.domain.reconciliation.

Cobre:
1.  parse_amount — valores formatados em pt_BR (vírgula decimal) não viram inteiros.
2.  normalize_soma_description — remove sufixo Nxxx do SOMA.
3.  Fluxo OK completo (trilogia T_EXTRATO→CONTAORDEM→SOMA).
4.  Não encontrado na CONTAORDEM.
5.  Não encontrado no SOMA.
6.  Duplicação de DOC. SOMA.
7.  EXT0000002976 não usa DOC. SOMA da origem quando CONTAORDEM não tem numérico.
"""
from __future__ import annotations

from decimal import Decimal

from soma_app.domain.reconciliation import (
    ReconciliationCategory,
    normalize_soma_description,
    parse_amount,
    reconcile_t_extrato,
)

# ---------------------------------------------------------------------------
# 1. parse_amount — preservação de valores pt_BR
# ---------------------------------------------------------------------------


class TestParseAmount:
    """Garante que valores formatados com vírgula decimal são parseados corretamente."""

    def test_230_virgula(self):
        """230,00 não vira 23000."""
        assert parse_amount("230,00") == Decimal("230.00")

    def test_1_virgula(self):
        """1,00 não vira 100."""
        assert parse_amount("1,00") == Decimal("1.00")

    def test_0_04(self):
        """0,04 permanece 0.04."""
        assert parse_amount("0,04") == Decimal("0.04")

    def test_50_virgula(self):
        """50,00 não vira 5000."""
        assert parse_amount("50,00") == Decimal("50.00")

    def test_125_virgula(self):
        """125,00 não vira 12500."""
        assert parse_amount("125,00") == Decimal("125.00")

    def test_negativo(self):
        """-289,21 é parseado corretamente."""
        assert parse_amount("-289,21") == Decimal("-289.21")

    def test_negativo_parens(self):
        """(289,21) é tratado como negativo."""
        assert parse_amount("(289,21)") == Decimal("-289.21")

    def test_ponto_milhar_virgula_decimal(self):
        """1.266,00 com separador de milhar ponto e decimal vírgula."""
        assert parse_amount("1.266,00") == Decimal("1266.00")

    def test_inteiro(self):
        """Inteiro 230 (sem vírgula) continua sendo 230."""
        assert parse_amount(230) == Decimal("230")

    def test_zero(self):
        assert parse_amount("0,00") == Decimal("0")

    def test_none(self):
        assert parse_amount(None) is None

    def test_string_vazia(self):
        assert parse_amount("") is None


# ---------------------------------------------------------------------------
# 2. normalize_soma_description
# ---------------------------------------------------------------------------


class TestNormalizeSomaDescription:
    """Remove sufixo Nxxx acrescentado pelo SOMA; não altera outros textos."""

    def test_despesas_bancarias_n004(self):
        result = normalize_soma_description("DESPESAS BANCÁRIAS N004")
        assert result == "DESPESAS BANCARIAS"

    def test_dizimos_n002(self):
        result = normalize_soma_description(
            "DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA) N002"
        )
        assert result == "DIZIMOS E OFERTAS (TRANSFERENCIA BANCARIA)"

    def test_n001(self):
        result = normalize_soma_description("PAG. CARTAO BUSTRADE N001")
        assert result == "PAG. CARTAO BUSTRADE"

    def test_sem_sufixo(self):
        """Descrição sem sufixo Nxxx não é alterada."""
        result = normalize_soma_description("DESPESAS BANCÁRIAS")
        assert result == "DESPESAS BANCARIAS"

    def test_numero_no_meio_preservado(self):
        """Número no meio do texto não é removido."""
        result = normalize_soma_description("LOTE 3 PAGAMENTO")
        assert result == "LOTE 3 PAGAMENTO"

    def test_ref_com_numeros_preservada(self):
        """Referência com números legítimos não é truncada."""
        result = normalize_soma_description("REF 2024/001 TAXA")
        assert result == "REF 2024/001 TAXA"

    def test_none(self):
        result = normalize_soma_description(None)
        assert result == ""


# ---------------------------------------------------------------------------
# 3. status_label — formatação do texto para a coluna STATUS
# ---------------------------------------------------------------------------


class TestStatusLabel:
    """Garante que a propriedade status_label gera exatamente as mensagens pedidas."""

    def test_confirmado(self):
        sources = [_source("EXT001", "DÍZIMOS E OFERTAS", "230,00")]
        contas = [_conta("EXT001", "DÍZIMOS E OFERTAS", "230,00", "5161859", descricao_soma="DÍZIMOS E OFERTAS")]
        somas = [_soma("5161859", "DÍZIMOS E OFERTAS", "230,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.status_label == "Confirmado"

    def test_nao_encontrado_contaordem(self):
        sources = [_source("EXT999", "TESTE", "10,00")]
        report = _run(sources, [], [])
        item = report.items[0]
        assert item.status_label == "Não Encontrado"

    def test_divergencia_importancia(self):
        sources = [_source("EXT001", "DÍZIMOS E OFERTAS", "230,00")]
        contas = [_conta("EXT001", "DÍZIMOS E OFERTAS", "500,00", "5161859")]
        somas = [_soma("5161859", "DÍZIMOS E OFERTAS", "500,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.status_label == "Divergência - IMPORTÂNCIA"

    def test_divergencia_data_mov(self):
        sources = [_source("EXT001", "DÍZIMOS E OFERTAS", "230,00", data="01/04/2026")]
        contas = [_conta("EXT001", "DÍZIMOS E OFERTAS", "230,00", "5161859", data="15/04/2026")]
        somas = [_soma("5161859", "DÍZIMOS E OFERTAS", "230,00", data="15/04/2026")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.status_label == "Divergência - DATA MOV."


# ---------------------------------------------------------------------------
# Helpers para construção de registos de teste
# ---------------------------------------------------------------------------


def _source(id_interno: str, descricao: str, valor: str, data: str = "01/04/2026",
            tipo: str = "Crédito", doc_soma: str = "") -> dict:
    return {
        "ID_INTERNO": id_interno,
        "DESCRIÇÃO": descricao,
        "IMPORTÂNCIA": valor,
        "DATA MOV.": data,
        "TIPO": tipo,
        "DOC. SOMA": doc_soma,
    }


def _conta(id_interno: str, descricao: str, valor: str, doc_soma: str,
           data: str = "01/04/2026", tipo: str = "Crédito",
           descricao_soma: str = "", processo: str = "") -> dict:
    return {
        "ID_INTERNO": id_interno,
        "DESCRIÇÃO": descricao,
        "IMPORTÂNCIA": valor,
        "DATA MOV.": data,
        "TIPO": tipo,
        "DOC. SOMA": doc_soma,
        "DESCRIÇÃO SOMA": descricao_soma or descricao,
        "PROCESSO": processo,
    }


def _soma(codigo: str, descricao: str, valor: str, data: str = "01/04/2026",
          tipo: str = "Crédito", status: str = "PAGO", baixa: str = "SIM") -> dict:
    return {
        "CODIGO": codigo,
        "DESCRIÇÃO": descricao,
        "VALOR": valor,
        "PAGAMENTO": data,
        "TIPO": tipo,
        "STATUS": status,
        "BAIXA": baixa,
    }


def _run(sources, contas, somas, year=2026, month=4):
    return reconcile_t_extrato(
        sources,
        contas,
        somas,
        year=year,
        month=month,
        exclusions=(),
    )


# ---------------------------------------------------------------------------
# 3. Fluxo OK completo
# ---------------------------------------------------------------------------


class TestFluxoOK:
    def test_ok_simples(self):
        """Trilogia T_EXTRATO→CONTAORDEM→SOMA sem divergências."""
        sources = [_source("EXT001", "DÍZIMOS E OFERTAS", "230,00")]
        contas = [_conta("EXT001", "DÍZIMOS E OFERTAS", "230,00", "5161859",
                         descricao_soma="DÍZIMOS E OFERTAS")]
        somas = [_soma("5161859", "DÍZIMOS E OFERTAS", "230,00")]

        report = _run(sources, contas, somas)
        assert len(report.items) == 1
        item = report.items[0]
        assert item.category == ReconciliationCategory.OK
        assert not item.issues

    def test_ok_descricao_com_sufixo_nxxx(self):
        """SOMA com sufixo Nxxx na descrição: deve ser OK (não divergência)."""
        sources = [_source("EXT002", "DESPESAS BANCÁRIAS", "50,00")]
        contas = [_conta("EXT002", "DESPESAS BANCÁRIAS", "50,00", "9999001",
                         descricao_soma="DESPESAS BANCÁRIAS")]
        # SOMA acrescentou N002 ao final
        somas = [_soma("9999001", "DESPESAS BANCÁRIAS N002", "50,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.category == ReconciliationCategory.OK, (
            f"Esperado OK mas obteve {item.category}; issues={item.issues}"
        )

    def test_ok_dizimos_com_sufixo_n004(self):
        """DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA) N004 → não divergência."""
        sources = [_source("EXT003", "DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA)", "1.000,00")]
        contas = [_conta("EXT003", "DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA)",
                         "1.000,00", "8888001",
                         descricao_soma="DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA)")]
        somas = [_soma("8888001", "DÍZIMOS E OFERTAS (TRANSFERENCIA BANCARIA) N004", "1.000,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.category == ReconciliationCategory.OK, (
            f"Esperado OK mas obteve {item.category}; issues={item.issues}"
        )

    def test_descricoes_realmente_diferentes_continuam_divergentes(self):
        """Descrições genuinamente diferentes devem continuar a gerar divergência."""
        sources = [_source("EXT004", "ALUGUER SALA", "300,00")]
        contas = [_conta("EXT004", "ALUGUER SALA", "300,00", "7777001",
                         descricao_soma="ALUGUER SALA")]
        # SOMA tem descrição completamente diferente
        somas = [_soma("7777001", "DÍZIMOS E OFERTAS N001", "300,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        # Deve haver pelo menos uma issue de descrição
        codes = {i.code for i in item.issues}
        assert "DESCRICAO_CONTA_SOMA" in codes, (
            f"Esperada issue DESCRICAO_CONTA_SOMA; issues={item.issues}"
        )


# ---------------------------------------------------------------------------
# 4. Não encontrado na CONTAORDEM
# ---------------------------------------------------------------------------


class TestNaoEncontradoContaordem:
    def test_id_nao_existe(self):
        sources = [_source("EXT999", "TESTE", "10,00")]
        contas = [_conta("EXT_OUTRO", "OUTRO", "10,00", "1111")]
        somas = [_soma("1111", "OUTRO", "10,00")]

        report = _run(sources, contas, somas)
        assert len(report.items) == 1
        item = report.items[0]
        assert item.category == ReconciliationCategory.NAO_ENCONTRADO_CONTAORDEM
        codes = {i.code for i in item.issues}
        assert "CONTAORDEM_NAO_ENCONTRADA" in codes


# ---------------------------------------------------------------------------
# 5. Não encontrado no SOMA
# ---------------------------------------------------------------------------


class TestNaoEncontradoSoma:
    def test_doc_soma_nao_existe_na_soma(self):
        sources = [_source("EXT100", "DEPÓSITO", "500,00")]
        contas = [_conta("EXT100", "DEPÓSITO", "500,00", "9876543")]
        somas = []  # SOMA vazio

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA
        codes = {i.code for i in item.issues}
        assert "SOMA_NAO_ENCONTRADO" in codes

    def test_contaordem_sem_doc_soma_numerico(self):
        """CONTAORDEM com DOC. SOMA não numérico → NAO_ENCONTRADO_SOMA."""
        sources = [_source("EXT101", "DEPÓSITO", "500,00")]
        contas = [_conta("EXT101", "DEPÓSITO", "500,00", "PGTO CARTÃO")]
        somas = [_soma("5161859", "OUTRA COISA", "230,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        assert item.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA
        codes = {i.code for i in item.issues}
        assert "CONTAORDEM_SEM_DOC_SOMA" in codes


# ---------------------------------------------------------------------------
# 6. Duplicação de DOC. SOMA
# ---------------------------------------------------------------------------


class TestDuplicacao:
    def test_dup_doc_contaordem(self):
        """Dois registos na CONTAORDEM com o mesmo DOC. SOMA e IDs diferentes."""
        sources = [_source("EXT200", "COMPRA A", "100,00")]
        contas = [
            _conta("EXT200", "COMPRA A", "100,00", "2222222"),
            _conta("EXT201", "COMPRA B", "100,00", "2222222"),  # mesmo DOC. SOMA
        ]
        somas = [_soma("2222222", "COMPRA A", "100,00")]

        report = _run(sources, contas, somas)
        item = report.items[0]
        codes = {i.code for i in item.issues}
        assert "DUP_DOC_CONTAORDEM" in codes
        assert item.category == ReconciliationCategory.DUPLICADO

    def test_dup_codigo_soma(self):
        """CODIGO duplicado na sheet SOMA."""
        sources = [_source("EXT210", "TAXA", "75,00")]
        contas = [_conta("EXT210", "TAXA", "75,00", "3333333")]
        somas = [
            _soma("3333333", "TAXA", "75,00"),
            _soma("3333333", "TAXA DUPLICADA", "75,00"),
        ]

        report = _run(sources, contas, somas)
        item = report.items[0]
        codes = {i.code for i in item.issues}
        assert "DUP_CODIGO_SOMA" in codes


# ---------------------------------------------------------------------------
# 7. EXT0000002976 — não usa DOC. SOMA da origem quando CONTAORDEM não tem numérico
# ---------------------------------------------------------------------------


class TestExt0000002976:
    """
    EXT0000002976:
    - Origem T_EXTRATO tem DOC. SOMA = 5161859 (pertence a outro lançamento).
    - CONTAORDEM tem DOC. SOMA = "PGTO CARTÃO" (não numérico).
    - O código NÃO deve usar 5161859 da origem para localizar o SOMA.
    - Resultado esperado: NAO_ENCONTRADO_SOMA (CONTAORDEM_SEM_DOC_SOMA).
    """

    def test_nao_usa_doc_soma_da_origem(self):
        sources = [
            _source("EXT0000002976", "PAG. CARTAO BUSTRADE", "-289,21",
                    data="01/04/2026", tipo="Cartão", doc_soma="5161859"),
        ]
        contas = [
            _conta("EXT0000002976", "PAG. CARTAO BUSTRADE", "289,21",
                   doc_soma="PGTO CARTÃO", tipo="Cartão"),
        ]
        # SOMA tem o código 5161859 mas pertence a OUTRO lançamento
        somas = [
            _soma("5161859", "DÍZIMOS E OFERTAS N001", "230,00"),
        ]

        report = _run(sources, contas, somas)
        assert len(report.items) == 1
        item = report.items[0]

        # Não deve estar conciliado com o SOMA
        assert item.soma_rows == [], (
            f"EXT0000002976 não deve ter soma_rows; obteve {item.soma_rows}"
        )
        # Categoria esperada: NAO_ENCONTRADO_SOMA
        assert item.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA, (
            f"Esperado NAO_ENCONTRADO_SOMA mas obteve {item.category}"
        )
        codes = {i.code for i in item.issues}
        assert "CONTAORDEM_SEM_DOC_SOMA" in codes, (
            f"Esperado CONTAORDEM_SEM_DOC_SOMA; issues={item.issues}"
        )

    def test_doc_soma_5161859_nao_contaminado(self):
        """O código 5161859 deve estar disponível para EXT0000002970 (outro lançamento)."""
        sources = [
            _source("EXT0000002970", "DÍZIMOS E OFERTAS", "230,00",
                    data="01/04/2026", tipo="Crédito", doc_soma="5161859"),
            _source("EXT0000002976", "PAG. CARTAO BUSTRADE", "-289,21",
                    data="01/04/2026", tipo="Cartão", doc_soma="5161859"),
        ]
        contas = [
            _conta("EXT0000002970", "DÍZIMOS E OFERTAS", "230,00",
                   doc_soma="5161859", tipo="Crédito",
                   descricao_soma="DÍZIMOS E OFERTAS"),
            _conta("EXT0000002976", "PAG. CARTAO BUSTRADE", "289,21",
                   doc_soma="PGTO CARTÃO", tipo="Cartão"),
        ]
        somas = [
            _soma("5161859", "DÍZIMOS E OFERTAS N001", "230,00"),
        ]

        report = _run(sources, contas, somas)
        by_id = {item.id_interno: item for item in report.items}

        # EXT0000002970 deve estar OK
        item_2970 = by_id.get("EXT0000002970")
        assert item_2970 is not None
        assert item_2970.category == ReconciliationCategory.OK, (
            f"EXT0000002970 esperado OK; issues={item_2970.issues}"
        )

        # EXT0000002976 deve estar NAO_ENCONTRADO_SOMA (não OK)
        item_2976 = by_id.get("EXT0000002976")
        assert item_2976 is not None
        assert item_2976.category == ReconciliationCategory.NAO_ENCONTRADO_SOMA, (
            f"EXT0000002976 esperado NAO_ENCONTRADO_SOMA; categoria={item_2976.category}"
        )
