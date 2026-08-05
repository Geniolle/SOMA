# 🔍 Análise do Servidor - Por que SOMA não está funcionando

**Data:** 2026-08-05  
**Status:** ❌ SOMA não está configurado para rodar no PM2

---

## 📊 Descobertas Principais

### 1️⃣ Processo `bot-igreja` (WhatsApp Bot)
```
Status: online (mas executando diferente projeto)
Script: /home/ubuntu/whatsapp-bot-reply/src/index.js
Tipo: Node.js (WhatsApp Bot)
Cwd: /home/ubuntu/whatsapp-bot-reply/
Restarts: 660 vezes (muito instável!)
Cron restart: Toda madrugada às 4AM
```

### 2️⃣ Projeto SOMA
```
Localização: ~/soma-automation/SOMA/
Tipo: Python (Selenium automation)
Status: ❌ NÃO configurado no PM2
Status: ❌ NÃO está rodando
Venv: ✅ Existe (.venv)
Main.py: ✅ Existe
.env: ✅ Deve existir em deploy/.env
```

---

## 🔴 Problemas Identificados

### Problema 1: SOMA não tem app PM2
- ❌ Nenhum arquivo `ecosystem.config.js` ou `*.config.js`
- ❌ SOMA não está registrado no PM2
- ❌ Não há forma automática de iniciar SOMA

### Problema 2: bot-igreja com 660 restarts
- ⚠️ Processo WhatsApp está muito instável
- ⚠️ Reinicia todo dia às 4AM automaticamente
- ⚠️ Pode estar consumindo recursos

### Problema 3: Confusão de processos
- ⚠️ SOMA e bot-igreja são projetos completamente diferentes
- ⚠️ Estão em diretórios diferentes
- ⚠️ Usam linguagens diferentes (Python vs Node.js)

---

## ✅ Checklist de Status

| Item | Status | Ação |
|------|--------|------|
| Venv Python | ✅ Existe | OK |
| main.py | ✅ Existe | OK |
| .env config | ❓ Desconhecido | Verificar |
| Git repo | ✅ Sincronizado | OK |
| Requirements | ❓ Desconhecido | Verificar |
| PM2 config | ❌ NÃO existe | **Criar** |
| PM2 app | ❌ NÃO existe | **Adicionar** |

---

## 🚀 Solução Recomendada

### Opção A: Criar arquivo ecosystem.config.js (Recomendado)
```bash
# Criar arquivo de configuração PM2
cat > ~/soma-automation/SOMA/ecosystem.config.js << 'EOF'
module.exports = {
  apps: [
    {
      name: "soma-automation",
      script: "main.py",
      interpreter: "/home/ubuntu/soma-automation/SOMA/.venv/bin/python",
      cwd: "/home/ubuntu/soma-automation/SOMA",
      instances: 1,
      exec_mode: "fork",
      env: {
        RUN_ENV: "production",
      },
      log_date_format: "YYYY-MM-DD HH:mm:ss Z",
      error_file: "logs/soma-error.log",
      out_file: "logs/soma-out.log",
      merge_logs: true,
      max_size: "100M",
      max_file: 14,
      autorestart: true,
      watch: false,
      ignore_watch: ["logs", "artifacts", ".venv", "node_modules"],
      min_uptime: "10s",
      max_restarts: 10,
      restart_delay: 4000,
    }
  ]
};
EOF

# Registrar no PM2
pm2 start ecosystem.config.js --name soma-automation
pm2 save
```

### Opção B: Iniciar manualmente
```bash
cd ~/soma-automation/SOMA
.venv/bin/python main.py
```

### Opção C: Usar deploy-server.sh (já criado)
```bash
bash deploy-server.sh
```

---

## 📝 Próximas Ações

1. **Verificar .env** - Confirmar que `deploy/.env` existe e está configurado
2. **Testar main.py** - Rodar manualmente para ver se funciona
3. **Criar ecosystem.config.js** - Para integração com PM2
4. **Registrar no PM2** - `pm2 start ecosystem.config.js`
5. **Monitorar** - `pm2 monit` ou `pm2 logs soma-automation`

---

## 🆘 Comandos Úteis

```bash
# Status
pm2 status

# Logs SOMA
pm2 logs soma-automation

# Restart SOMA
pm2 restart soma-automation

# Parar SOMA
pm2 stop soma-automation

# Iniciar SOMA
pm2 start ecosystem.config.js

# Remover SOMA de PM2
pm2 delete soma-automation

# Dashboard
pm2 monit
```

---

## ⚠️ Nota Importante

**bot-igreja** é um bot WhatsApp completamente diferente do SOMA:
- Localização diferente: `/home/ubuntu/whatsapp-bot-reply/`
- Linguagem diferente: Node.js (não Python)
- Propósito diferente: Bot de resposta WhatsApp (não automação web)

Não confundir os dois projetos!

---

**Status:** 🔴 SOMA requer configuração no PM2 para funcionar automaticamente
