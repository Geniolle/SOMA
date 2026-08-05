# 📋 Guia de Limpeza de Logs - SOMA

## 📊 Status Atual do Servidor

```
Diretório: ~/soma-automation/SOMA/logs/
Tamanho total: 52MB

Arquivos principais:
- soma-run.log: 42MB (arquivo ativo)
- soma-cron.log: 151KB
- logs antigos: ~10MB (soma_dev_*.log de maio, junho, julho)
```

---

## 🛠️ Ferramentas Disponíveis

### Opção 1: Script Python (Recomendado)

**Arquivo:** `manage_logs.py`

**Uso básico:**
```bash
cd ~/soma-automation/SOMA
python manage_logs.py
```

**Opções avançadas:**
```bash
# Manter últimos 14 dias (ao invés de 30)
python manage_logs.py --keep-days 14

# Rotacionar arquivo se > 500MB (ao invés de 1GB)
python manage_logs.py --max-size 500

# Usar diretório customizado
python manage_logs.py --dir /var/log/soma
```

**O que faz:**
✅ Deleta logs de entrada/saída (`soma_dev_*.log`) antigos  
✅ Deleta arquivos de log vazios  
✅ Rotaciona `soma-run.log` se > 1GB  
✅ Compacta arquivos rotacionados com gzip  
✅ Mostra resumo de espaço liberado  
✅ Lista arquivos de log disponíveis  

---

### Opção 2: Script Bash (Manual)

**Arquivo:** `cleanup_logs.sh`

**Uso básico:**
```bash
cd ~/soma-automation/SOMA
chmod +x cleanup_logs.sh
./cleanup_logs.sh
```

**Com parâmetros:**
```bash
# Manter últimos 14 dias
./cleanup_logs.sh logs 14

# Máximo 500MB para arquivo principal
./cleanup_logs.sh logs 30 500
```

---

## 🔄 Automação (Cron Job)

Para automatizar a limpeza de logs, adicione ao crontab:

```bash
# Editar crontab
crontab -e

# Adicionar uma destas linhas:

# Executar diariamente às 2 da manhã
0 2 * * * cd ~/soma-automation/SOMA && python manage_logs.py >> logs/cleanup.log 2>&1

# Executar a cada 6 horas
0 */6 * * * cd ~/soma-automation/SOMA && python manage_logs.py >> logs/cleanup.log 2>&1

# Executar todo domingo às 3 da manhã
0 3 * * 0 cd ~/soma-automation/SOMA && python manage_logs.py >> logs/cleanup.log 2>&1
```

---

## 📋 Exemplo de Saída

```
================================================================================
GERENCIADOR DE LOGS - SOMA
================================================================================

📊 ANTES DA LIMPEZA: 52.0MB

🧹 EXECUTANDO LIMPEZA...
  🗑️  Deletado: soma_dev_20260602_024546.log
  🗑️  Deletado: soma_dev_20260519_022048.log
  🗑️  Deletado: soma_dev_20260512_020630.log
  ... (10 logs deletados)
  ✅ 10 logs antigos removidos
  ✅ 0 logs vazios removidos
  ⚠️  soma-run.log muito grande: 42.0MB
  🔄 Rotacionando arquivo...
  ✅ Arquivo rotacionado para: soma-run-20260805_214530.log
  📦 Arquivo comprimido: soma-run-20260805_214530.log.gz
     Tamanho: 4.2MB
  ✅ Arquivo principal OK

📊 DEPOIS DA LIMPEZA: 15.0MB
💾 Espaço liberado: 37.0MB

📋 ARQUIVOS DE LOG (primeiros 15):
Nome                                               Tamanho       Data
────────────────────────────────────────────────────────────────────────────
soma-cron.log                                      151.0KB    2026-08-05 21:20
soma_dev_20260704_020315.log                       173.0KB    2026-07-04 02:03
soma_dev_20260710_022216.log                       174.0KB    2026-07-10 02:43
...

================================================================================
✅ LIMPEZA CONCLUÍDA!
================================================================================
```

---

## 🎯 Recomendações

| Cenário | Ação | Frequência |
|---------|------|-----------|
| Servidor em produção | Rodar `manage_logs.py` | Diariamente (cron) |
| Desenvolvimento local | Rodar manualmente | Conforme necessário |
| Após testes intensos | Rodar com `--keep-days 7` | Conforme necessário |
| Armazenamento crítico | Rodar com `--max-size 256` | A cada 12 horas |

---

## 📊 Métricas

**Economia de espaço típica:**
- Remover 10 logs antigos: ~5-10MB
- Compactar soma-run.log 42MB: ~4-5MB  
- **Total liberado:** ~10-15MB por execução

**Tempo de execução:** ~2-5 segundos

**Segurança:** Apenas deleta logs com >30 dias (configurável)

---

## ⚠️ Notas Importantes

1. **Backup:** Script não faz backup automático. Se necessário arquivar logs:
   ```bash
   tar -czf logs-backup-$(date +%Y%m%d).tar.gz logs/
   ```

2. **Monitoramento:** Verifique `cleanup.log` gerado pelo cron:
   ```bash
   tail -f logs/cleanup.log
   ```

3. **Recuperação:** Logs deletados não podem ser recuperados. Certifique-se de que os logs foram enviados para um sistema de logging centralizado se necessário.

---

## 🚀 Próximos Passos

1. ✅ Testar `manage_logs.py` manualmente
2. ✅ Configurar cron job se produção
3. ✅ Monitorar primeira execução
4. ✅ Ajustar parâmetros conforme necessário

---

**Status:** 🟢 Ferramentas disponíveis e prontas para usar
