#!/usr/bin/env python3
"""Simular como ficaria o quadro (sheet) com as informacoes."""

import json
from pathlib import Path

# Carregar resultado
result_file = Path("output/resultado.json")
data = json.loads(result_file.read_text(encoding="utf-8"))

movimentos = data["movimentos"][:10]  # Primeiros 10 para simulacao

print("\n" + "="*180)
print("SIMULACAO DE LAYOUT - SHEET CARTAO")
print("="*180)

# Cenário 1: Estado Atual (sem linha 5 vazia)
print("\n\n[CENARIO 1] ESTADO ATUAL (Rollback - sem deslocamento)")
print("-" * 180)
print(f"{'Ln':<4} {'ID_INTERNO':<15} {'Data Mov':<12} {'Data Valor':<12} {'Descricao':<30} {'Debito EUR':<12} {'Credito EUR':<12} {'Status':<10}")
print("-" * 180)

for i, mov in enumerate(movimentos, 1):
    linha = mov['line']
    desc = mov['descricao'][:28] if mov['descricao'] else "(vazio)"
    debito = mov['debito_eur'] if mov['debito_eur'] else "-"
    credito = mov['credito_eur'] if mov['credito_eur'] else "-"
    status = "OK" if mov['status'] == "VALIDO" else "REV"

    # Para cenário 1, usar índice natural 1, 2, 3...
    print(f"{i:<4} {'CAR' + str(i).zfill(10):<15} {mov['data_movimento']:<12} {mov['data_valor']:<12} {desc:<30} {debito:<12} {credito:<12} {status:<10}")

# Cenário 2: Com linha 5 vazia como cabeçalho
print("\n\n[CENARIO 2] COM LINHA 5 VAZIA (Desejo do usuario)")
print("-" * 180)
print(f"{'Ln':<4} {'ID_INTERNO':<15} {'Data Mov':<12} {'Data Valor':<12} {'Descricao':<30} {'Debito EUR':<12} {'Credito EUR':<12} {'Status':<10}")
print("-" * 180)

# Linha 5 vazia (cabeçalho)
print(f"{'5':<4} {'-':<15} {'':<12} {'':<12} {'[CABECALHO - VAZIO]':<30} {'':<12} {'':<12} {'--':<10}")

# Dados começam na linha 6
for i, mov in enumerate(movimentos, 1):
    linha_deslocada = i + 5  # Deslocar tudo em +5
    desc = mov['descricao'][:28] if mov['descricao'] else "(vazio)"
    debito = mov['debito_eur'] if mov['debito_eur'] else "-"
    credito = mov['credito_eur'] if mov['credito_eur'] else "-"
    status = "OK" if mov['status'] == "VALIDO" else "REV"
    id_interno = f"CAR{i:010d}"

    print(f"{linha_deslocada:<4} {id_interno:<15} {mov['data_movimento']:<12} {mov['data_valor']:<12} {desc:<30} {debito:<12} {credito:<12} {status:<10}")

# Cenário 3: Apenas Data Mov vazia na linha 5
print("\n\n[CENARIO 3] APENAS Data Mov VAZIA NA LINHA 5 (parcial)")
print("-" * 180)
print(f"{'Ln':<4} {'ID_INTERNO':<15} {'Data Mov':<12} {'Data Valor':<12} {'Descricao':<30} {'Debito EUR':<12} {'Credito EUR':<12} {'Status':<10}")
print("-" * 180)

# Linha 5 com Data Mov vazia mas resto preenchido
print(f"{'5':<4} {'CAR0000000001':<15} {'':<12} {'[HEADER]':<12} {'[CABECALHO]':<30} {'':<12} {'':<12} {'--':<10}")

# Dados naturais começam na linha 6
for i, mov in enumerate(movimentos, 1):
    linha_deslocada = i + 5
    desc = mov['descricao'][:28] if mov['descricao'] else "(vazio)"
    debito = mov['debito_eur'] if mov['debito_eur'] else "-"
    credito = mov['credito_eur'] if mov['credito_eur'] else "-"
    status = "OK" if mov['status'] == "VALIDO" else "REV"
    id_interno = f"CAR{(i+1):010d}"

    print(f"{linha_deslocada:<4} {id_interno:<15} {mov['data_movimento']:<12} {mov['data_valor']:<12} {desc:<30} {debito:<12} {credito:<12} {status:<10}")

print("\n" + "="*180)
print("LEGENDA")
print("="*180)
print("* Cenario 1: Estado atual (sem linhas vazias inseridas)")
print("* Cenario 2: Linha 5 completamente vazia, dados comecam linha 6, IDs sequenciais")
print("* Cenario 3: Linha 5 com Data Mov vazia apenas, resto preenchido com cabecalho")
print("* OK = VALIDO | REV = REVISAO")
print("="*180 + "\n")
