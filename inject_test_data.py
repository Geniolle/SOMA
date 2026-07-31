#!/usr/bin/env python3
"""
Script para injetar dados de teste na sheet CONTAORDEM do Google Sheets.
"""
from __future__ import annotations

import sys
from pathlib import Path

base_dir = Path(__file__).resolve().parent
src_dir = base_dir / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

import os
from dotenv import load_dotenv

load_dotenv(base_dir / "deploy" / ".env")

from soma_app.infra.sheets_client import SheetsClient


def inject_test_data():
    """Injeta uma linha de teste na sheet CONTAORDEM."""
    print("[*] Inicializando SheetsClient...")

    class MockSettings:
        pass

    settings = MockSettings()
    sheets = SheetsClient(settings)

    ws_name = os.getenv("SHEET_CONTAORDEM", "CONTAORDEM")
    print(f"[*] Sheet alvo: {ws_name}")

    try:
        # Lê dados atuais para entender a estrutura
        records = sheets.get_all_records_raw(ws_name)
        print(f"[*] Registros atuais: {len(records)}")

        if records:
            headers = list(records[0].keys())
            print(f"[*] Headers: {headers}")
        else:
            print("[!] Sheet vazia ou sem headers")
            return False

        # Monta linha de teste com os headers existentes (use lowercase + espacos)
        header_lower = [h.lower() for h in headers]
        test_row = {}
        for i, h_lower in enumerate(header_lower):
            h_orig = headers[i]
            if "tipo" in h_lower:
                test_row[h_orig] = "Saida"
            elif "status" in h_lower:
                test_row[h_orig] = "Novo"
            elif "plano" in h_lower:
                test_row[h_orig] = "1.3.2"
            elif "centro" in h_lower or "custo" in h_lower:
                test_row[h_orig] = "CC001"
            elif "descri" in h_lower and "soma" in h_lower:
                test_row[h_orig] = "Teste AUTO - TRF Bancaria"
            elif "import" in h_lower or "valor" in h_lower:
                test_row[h_orig] = "100.50"
            elif "forma" in h_lower or "pagamento" in h_lower:
                test_row[h_orig] = "Transferencia Bancaria"
            elif "caixa" in h_lower and "saida" not in h_lower:
                test_row[h_orig] = "BANCO ITAU"
            elif "id_interno" in h_lower or "interno" in h_lower:
                test_row[h_orig] = "TEST-001"
            else:
                test_row[h_orig] = ""

        # Converte para lista ordenada pelos headers
        values = [test_row.get(h, "") for h in headers]
        print(f"[*] Valores a inserir: {len(values)} colunas")

        # Encontra proxima linha vazia
        row_num = len(records) + 2  # +1 para header, +1 para nova linha

        # Usa batch_update para inserir
        ranges = [
            (f"A{row_num}:Z{row_num}", [values])
        ]
        sheets.batch_update(ws_name, ranges)

        print("[OK] Linha de teste adicionada na linha %d" % row_num)
        return True

    except Exception as e:
        print(f"[!] Erro: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = inject_test_data()
    sys.exit(0 if success else 1)
