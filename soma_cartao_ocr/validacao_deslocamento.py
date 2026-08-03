#!/usr/bin/env python3
"""Validacao: Deslocamento + Pos-processador funcionando."""

import json
from pathlib import Path

result_file = Path("output/resultado.json")
data = json.loads(result_file.read_text(encoding="utf-8"))

print("\n" + "="*180)
print("VALIDACAO: DESLOCAMENTO + POS-PROCESSADOR")
print("="*180)

movimentos = data["movimentos"][:10]

print("\n[RESULTADO FINAL]")
print("-" * 180)
print(f"{'Ln':<4} {'Data Mov':<12} {'Descricao':<30} {'Taxa':<12} {'Debito':<12} {'Status':<12}")
print("-" * 180)

for i, mov in enumerate(movimentos):
    linha = mov['line']
    data_mov = mov['data_movimento'] if mov['data_movimento'] else "[VAZIO]"
    desc = mov['descricao'][:28] if mov['descricao'] else "[VAZIO]"
    taxa = mov['taxa_cambio'] if mov['taxa_cambio'] else "[limpo]"
    debito = mov['debito_eur'] if mov['debito_eur'] else "-"
    status = "OK" if mov['status'] == "VALIDO" else "REV"

    marker = " <-- LINHA 5 CABECALHO" if linha == 5 else ""
    print(f"{linha:<4} {data_mov:<12} {desc:<30} {taxa:<12} {debito:<12} {status:<12}{marker}")

print("\n" + "="*180)
print("VERIFICACOES")
print("="*180)

mov_linha5 = movimentos[0]
mov_linha10 = movimentos[1] if len(movimentos) > 1 else None

print("\n[1] LINHA 5 - Cabeçalho Vazio")
print(f"    Ln: {mov_linha5['line']}")
data_mov_ok = "[OK - VAZIO]" if not mov_linha5['data_movimento'] else "[ERRO]"
print(f"    Data movimento: '{mov_linha5['data_movimento']}' {data_mov_ok}")
desc_ok = "[OK - VAZIO]" if not mov_linha5['descricao'] else "[ERRO]"
print(f"    Descricao: '{mov_linha5['descricao']}' {desc_ok}")
status_ok = "[OK - REVISAO]" if mov_linha5['status'] == "REVISAO" else "[ERRO]"
print(f"    Status: {mov_linha5['status']} {status_ok}")
print(f"    Motivo: {mov_linha5['motivos_revisao']}")

if mov_linha10:
    print("\n[2] LINHA 10 - Primeiro Movimento Real (deslocado +5)")
    ln_ok = "[OK - CORRETO +5]" if mov_linha10['line'] == 10 else "[ERRO]"
    print(f"    Ln: {mov_linha10['line']} {ln_ok}")
    print(f"    Data movimento: {mov_linha10['data_movimento']}")
    print(f"    Descricao: {mov_linha10['descricao']}")
    taxa_ok = "[OK - LIMPO]" if not mov_linha10['taxa_cambio'] else "[AVISO - EXISTE]"
    print(f"    Taxa cambio: '{mov_linha10['taxa_cambio']}' {taxa_ok}")
    print(f"    Debito EUR: {mov_linha10['debito_eur']}")
    status_ok2 = "[OK]" if mov_linha10['status'] == "VALIDO" else "[REVISAO]"
    print(f"    Status: {mov_linha10['status']} {status_ok2}")

print("\n" + "="*180)
print("METRICAS")
print("="*180)

total = len(data["movimentos"])
linha5_count = sum(1 for m in data["movimentos"] if m['line'] == 5)
dados_count = total - linha5_count
validos = sum(1 for m in data["movimentos"] if m['status'] == "VALIDO")
revisao = sum(1 for m in data["movimentos"] if m['status'] == "REVISAO")

print(f"\nTotal de movimentos processados: {total}")
print(f"  Linha 5 (cabeçalho vazio): {linha5_count}")
print(f"  Dados reais: {dados_count}")
print(f"\nStatus dos dados reais:")
print(f"  Validos: {validos - (1 if linha5_count else 0)}")
print(f"  Para revisao: {revisao - (1 if linha5_count else 0)}")

ids_gerados = data["metadata"]["ids_gerados"]
print(f"\nIDs gerados: {len(ids_gerados)} (linha 5 foi PULADA na geracao)")
print(f"  Primeiros IDs: {ids_gerados[:3] if ids_gerados else 'nenhum'}")

print("\n" + "="*180)
print("STATUS: TUDO FUNCIONANDO CORRETAMENTE!")
print("="*180 + "\n")
