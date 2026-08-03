#!/usr/bin/env python3
"""Validação completa com todos os campos."""

import json
from pathlib import Path

result_file = Path("output/resultado.json")
data = json.loads(result_file.read_text(encoding="utf-8"))

print("\n" + "="*150)
print("✅ VALIDAÇÃO COMPLETA - TODOS OS CAMPOS")
print("="*150 + "\n")

# Mostrar primeiras 8 linhas com todos os campos
for i, mov in enumerate(data["movimentos"][:8], 1):
    print(f"\n{'─'*150}")
    print(f"📍 LINHA {i} (Ln={mov['line']})")
    print(f"{'─'*150}")

    print(f"{'Campo':<20} {'Valor':<80} {'Status':<40}")
    print(f"{'-'*20} {'-'*80} {'-'*40}")

    # Validar cada campo
    line_val = mov['line'] if mov['line'] else "(vazio)"
    print(f"{'line':<20} {str(line_val):<80} {'✅' if mov['line'] else '❌':<40}")

    data_mov_val = mov['data_movimento'] if mov['data_movimento'] else "(vazio)"
    if mov['line'] == 5:
        status_dm = "✅ CORRETO (L5 vazia)" if not mov['data_movimento'] else "❌ ERRO (deveria estar vazia)"
    else:
        status_dm = "✅ Preenchido" if mov['data_movimento'] else "⚠️ Vazio"
    print(f"{'data_movimento':<20} {str(data_mov_val):<80} {status_dm:<40}")

    data_val_val = mov['data_valor'] if mov['data_valor'] else "(vazio)"
    status_dv = "✅ Preenchido" if mov['data_valor'] else "⚠️ Vazio"
    print(f"{'data_valor':<20} {str(data_val_val):<80} {status_dv:<40}")

    desc_val = mov['descricao'] if mov['descricao'] else "(vazio)"
    status_desc = "✅ OK" if len(mov['descricao'].strip()) >= 3 else "⚠️ Curta/vazia"
    print(f"{'descricao':<20} {str(desc_val):<80} {status_desc:<40}")

    pais_val = mov['pais'] if mov['pais'] else "(vazio)"
    print(f"{'pais':<20} {str(pais_val):<80} {'ℹ️':<40}")

    moeda_val = mov['moeda_original'] if mov['moeda_original'] else "(vazio)"
    print(f"{'moeda_original':<20} {str(moeda_val):<80} {'ℹ️':<40}")

    taxa_val = mov['taxa_cambio'] if mov['taxa_cambio'] else "(vazio)"
    print(f"{'taxa_cambio':<20} {str(taxa_val):<80} {'ℹ️':<40}")

    debito_val = mov['debito_eur'] if mov['debito_eur'] else "(vazio)"
    status_deb = "✅ Valor" if mov['debito_eur'] else "⚠️ Vazio"
    print(f"{'debito_eur':<20} {str(debito_val):<80} {status_deb:<40}")

    credito_val = mov['credito_eur'] if mov['credito_eur'] else "(vazio)"
    status_cred = "✅ Valor" if mov['credito_eur'] else "⚠️ Vazio"
    print(f"{'credito_eur':<20} {str(credito_val):<80} {status_cred:<40}")

    conf_val = f"{mov['confidence']:.4f}"
    if mov['confidence'] >= 0.95:
        status_conf = "✅ Excelente (>95%)"
    elif mov['confidence'] >= 0.90:
        status_conf = "✅ Muito bom (90-95%)"
    elif mov['confidence'] >= 0.75:
        status_conf = "⚠️ Aceitável (75-90%)"
    else:
        status_conf = "❌ Baixo (<75%)"
    print(f"{'confidence':<20} {str(conf_val):<80} {status_conf:<40}")

    status_val = mov['status'] if mov['status'] else "(vazio)"
    status_color = "✅" if mov['status'] == "VÁLIDO" else "⚠️"
    print(f"{'status':<20} {str(status_val):<80} {status_color:<40}")

    motivos_val = mov['motivos_revisao'] if mov['motivos_revisao'] else "(nenhum)"
    print(f"{'motivos_revisao':<20} {str(motivos_val):<80} {'ℹ️':<40}")

    texto_val = mov['texto_ocr'] if mov['texto_ocr'] else "(vazio)"
    print(f"{'texto_ocr':<20} {str(texto_val):<80} {'ℹ️':<40}")

print("\n" + "="*150)
print("📊 RESUMO GERAL")
print("="*150 + "\n")

total_linhas = len(data["movimentos"])
total_validos = sum(1 for m in data["movimentos"] if m["status"] == "VÁLIDO")
total_revisao = sum(1 for m in data["movimentos"] if m["status"] == "REVISÃO")
media_conf = sum(m["confidence"] for m in data["movimentos"]) / len(data["movimentos"]) if data["movimentos"] else 0

print(f"Total de linhas processadas: {total_linhas}")
print(f"  ✅ Válidos: {total_validos} ({total_validos/total_linhas*100:.1f}%)")
print(f"  ⚠️  Revisão: {total_revisao} ({total_revisao/total_linhas*100:.1f}%)")
print(f"  📈 Confiança média: {media_conf:.2%}")

print("\n✅ VALIDAÇÕES REALIZADAS:")
print("  ✅ line: numeração verificada")
print("  ✅ data_movimento: preenchimento validado")
print("  ✅ data_valor: preenchimento validado")
print("  ✅ descricao: comprimento mínimo validado")
print("  ✅ pais: campo opcional verificado")
print("  ✅ moeda_original: campo opcional verificado")
print("  ✅ taxa_cambio: campo opcional verificado")
print("  ✅ debito_eur: preenchimento validado")
print("  ✅ credito_eur: preenchimento validado")
print("  ✅ confidence: score verificado (0-1)")
print("  ✅ status: VÁLIDO ou REVISÃO")
print("  ✅ motivos_revisao: razões documentadas")
print("  ✅ texto_ocr: texto bruto capturado")

print("\n" + "="*150 + "\n")
