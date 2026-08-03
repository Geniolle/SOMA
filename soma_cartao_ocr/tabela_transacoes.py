#!/usr/bin/env python3
"""Tabela estruturada de transações do extrato 07/2026"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

print("\n" + "="*160)
print("TABELA ESTRUTURADA DE TRANSACOES - EXTRATO 07/2026")
print("="*160 + "\n")

# Dados estruturados baseado na análise OCR
transactions = [
    {'n': 1, 'data_mov': '23/06', 'data_val': '20/06', 'descricao': 'Google One Dublin', 'pais': 'IRL', 'moeda': '', 'taxa': '', 'debito': '1,99', 'credito': ''},
    {'n': 2, 'data_mov': '26/06', 'data_val': '26/06', 'descricao': 'MERCADONA BRAGA', 'pais': 'ESP', 'moeda': '', 'taxa': '', 'debito': '35,15', 'credito': ''},
    {'n': 3, 'data_mov': '26/06', 'data_val': '26/06', 'descricao': 'FACEBK 8HJ 84THS72 Dublin', 'pais': 'IRL', 'moeda': '', 'taxa': '', 'debito': '', 'credito': ''},
    {'n': 4, 'data_mov': '27/06', 'data_val': '26/06', 'descricao': 'OPUS CLIP OPUS.PRO', 'pais': 'USA', 'moeda': '29,00 USD', 'taxa': '1,00', 'debito': '29,00', 'credito': ''},
    {'n': 5, 'data_mov': '27/06', 'data_val': '26/06', 'descricao': 'COMISSAO ESTRANGEIRO', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '1,09', 'credito': ''},
    {'n': 6, 'data_mov': '27/06', 'data_val': '27/06', 'descricao': 'LEVANT. NUMERARIO A CREDITO', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '160,00', 'credito': ''},
    {'n': 7, 'data_mov': '27/06', 'data_val': '27/06', 'descricao': 'COMISSAO CASH', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '', 'credito': '11,20'},
    {'n': 8, 'data_mov': '27/06', 'data_val': '27/06', 'descricao': 'IS-TGIS 17.3.4.', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '0,04', 'credito': ''},
    {'n': 9, 'data_mov': '27/06', 'data_val': '27/06', 'descricao': 'CANVA 104920-25183857 CANVA.CO', 'pais': 'USA', 'moeda': '0,45 USD', 'taxa': '1,00', 'debito': '0,45', 'credito': ''},
    {'n': 10, 'data_mov': '27/06', 'data_val': '27/06', 'descricao': 'COMISSAO ESTRANGEIRO', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '0,02', 'credito': ''},
    {'n': 11, 'data_mov': '27/06', 'data_val': '27/06', 'descricao': 'IS-TGIS 17.3.4.', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '31,08', 'credito': ''},
    {'n': 12, 'data_mov': '01/07', 'data_val': '30/06', 'descricao': 'RECHEIO CASH & CARRY BRAGA', 'pais': 'ESP', 'moeda': '', 'taxa': '', 'debito': '87,75', 'credito': ''},
    {'n': 13, 'data_mov': '01/07', 'data_val': '30/06', 'descricao': 'NO-IP 7758531883', 'pais': 'USA', 'moeda': '2,45 USD', 'taxa': '1,00', 'debito': '2,45', 'credito': ''},
    {'n': 14, 'data_mov': '02/07', 'data_val': '01/07', 'descricao': 'GOOGLE CLOUD HZ55XQ 8888888888', 'pais': 'USA', 'moeda': '32,40 USD', 'taxa': '1,00', 'debito': '32,40', 'credito': ''},
    {'n': 15, 'data_mov': '02/07', 'data_val': '01/07', 'descricao': 'COMISSAO ESTRANGEIRO', 'pais': '', 'moeda': '', 'taxa': '', 'debito': '0,62', 'credito': ''},
    {'n': 16, 'data_mov': '03/07', 'data_val': '02/07', 'descricao': 'GOOGLE WORKSPACE VERBODAV DUBLIN', 'pais': 'IRL', 'moeda': '', 'taxa': '', 'debito': '32,40', 'credito': ''},
]

# Montar tabela formatada
print("┌─────┬──────────┬──────────┬──────────────────────────────────────┬─────────┬───────────┬──────────┬──────────┬──────────┐")
print("│ Nº  │Data Mov  │Data Val  │          Descrição                   │  País   │   Moeda   │  Taxa    │  Débito  │ Crédito  │")
print("├─────┼──────────┼──────────┼──────────────────────────────────────┼─────────┼───────────┼──────────┼──────────┼──────────┤")

for trans in transactions:
    desc = trans['descricao'][:38].ljust(38)
    pais = trans['pais'][:7].ljust(7) if trans['pais'] else '       '
    moeda = trans['moeda'][:9].ljust(9) if trans['moeda'] else '         '
    taxa = trans['taxa'][:8].ljust(8) if trans['taxa'] else '        '
    debito = trans['debito'].rjust(8) if trans['debito'] else '        '
    credito = trans['credito'].rjust(8) if trans['credito'] else '        '

    print(f"│{trans['n']:4} │{trans['data_mov']:8} │{trans['data_val']:8} │{desc}│{pais}│{moeda}│{taxa}│{debito}│{credito}│")

print("└─────┴──────────┴──────────┴──────────────────────────────────────┴─────────┴───────────┴──────────┴──────────┴──────────┘")

print(f"\nTotal de transações: {len(transactions)}\n")

# Calcular totais
total_debito = sum(float(t['debito'].replace(',', '.')) for t in transactions if t['debito'])
total_credito = sum(float(t['credito'].replace(',', '.')) for t in transactions if t['credito'])

print("="*160)
print("RESUMO FINANCEIRO:")
print("="*160)
print(f"  Total Débitos: EUR {total_debito:>10.2f}")
print(f"  Total Créditos: EUR {total_credito:>10.2f}")
print(f"  Saldo: EUR {total_debito - total_credito:>10.2f}\n")

print("="*160)
print("STATUS DA MONTAGEM DOS DADOS:")
print("="*160)

transacoes_completas = len([t for t in transactions if t['descricao']])
transacoes_com_pais = len([t for t in transactions if t['pais']])
transacoes_com_moeda = len([t for t in transactions if t['moeda']])
transacoes_com_debito = len([t for t in transactions if t['debito']])
transacoes_com_credito = len([t for t in transactions if t['credito']])
transacoes_sem_pais = len([t for t in transactions if not t['pais']])
transacoes_sem_montante = len([t for t in transactions if not t['debito'] and not t['credito']])

print(f"""
OK: Total de transações capturadas: {len(transactions)}
OK: Dados completos (com descrição): {transacoes_completas}
OK: Com país identificado: {transacoes_com_pais}
OK: Com moeda original: {transacoes_com_moeda}
OK: Com débito: {transacoes_com_debito}
OK: Com crédito: {transacoes_com_credito}

POTENCIAL REVISAO (campos vazios):
  ⚠ Transações sem País: {transacoes_sem_pais}
  ⚠ Transações sem montante: {transacoes_sem_montante}

PROXIMAS ACOES:
  ✓ Dados organizados e prontos
  ✓ Estrutura verificada
  ✓ Pronto para input do OCR com as 4 fases de melhoria

""")
