#!/usr/bin/env python3
"""Mostrar sequência de linhas do resultado."""

import json
from pathlib import Path

result_file = Path("output/resultado.json")
data = json.loads(result_file.read_text(encoding="utf-8"))

print("\n" + "="*80)
print("📊 SEQUÊNCIA DE LINHAS NO RESULTADO JSON")
print("="*80 + "\n")

print(f"{'Ln':<4} | {'Data Mov':<10} | {'Descrição':<40}")
print("-"*80)

for mov in data["movimentos"]:
    ln = mov["line"]
    data_mov = mov["data_movimento"] if mov["data_movimento"] else "(vazio)"
    descricao = mov["descricao"][:37] if mov["descricao"] else "(vazio)"

    print(f"{ln:<4} | {data_mov:<10} | {descricao:<40}")

print("-"*80)
print(f"\nTotal de movimentos: {len(data['movimentos'])}\n")

# Análise
print("✅ ANÁLISE:\n")
linhas = [mov["line"] for mov in data["movimentos"]]
print(f"Primeira linha: {linhas[0]}")
print(f"Última linha: {linhas[-1]}")
print(f"\nSequência esperada:")
print(f"  • Linha 5: vazio (cabeçalho)")
print(f"  • Linha 6: primeira transação (23/06)")
print(f"  • Linha 7: segunda transação (26/06 MERCADONA)")
print(f"  • ...\n")

# Verificar deslocamento
print("🔍 VERIFICAÇÃO DO DESLOCAMENTO:\n")
if linhas[0] == 5:
    print("✅ Linha 5 é cabeçalho (vazio) - CORRETO")
else:
    print(f"❌ Linha 5 esperada, mas encontrada linha {linhas[0]}")

if len(linhas) > 1:
    primeira_transacao_linha = linhas[1]
    print(f"   Primeira transação está na linha {primeira_transacao_linha}")
    if primeira_transacao_linha == 6:
        print("   ✅ Primeira transação na linha 6 - CORRETO")
    else:
        print(f"   ❌ Esperada linha 6, encontrada {primeira_transacao_linha}")

print("\n" + "="*80 + "\n")
