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

### Problema 1: Chrome não instalado no servidor ❌ CRÍTICO
```
selenium.common.exceptions.SessionNotCreatedException
Message: session not created
from chrome not reachable
```
- ❌ Chrome não está instalado
- ❌ ChromeDriver não consegue ser inicializado
- ❌ SOMA depende de Chrome para funcionar

**Solução:** Instalar Google Chrome + ChromeDriver

### Problema 2: SOMA não tem app PM2 ✅ RESOLVIDO
- ✅ Arquivo `ecosystem.config.js` criado
- ✅ SOMA agora registrado no PM2 (soma-automation)
- ✅ Forma automática de iniciar SOMA pronta

### Problema 3: bot-igreja com 660 restarts
- ⚠️ Processo WhatsApp está muito instável
- ⚠️ Reinicia todo dia às 4AM automaticamente
- ⚠️ Pode estar consumindo recursos

### Problema 4: Confusão de processos
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

## 🚀 Solução - Instalar Chrome

### 1️⃣ Instalar Google Chrome (Ubuntu)
```bash
ssh -i chave.key ubuntu@132.145.57.133 << 'SSH_EOF'

echo "Instalando Google Chrome..."
curl -fsSL https://dl-ssl.google.com/linux/linux_signing_key.pub | sudo apt-key add -
sudo sh -c 'echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list'
sudo apt-get update
sudo apt-get install -y google-chrome-stable

echo "Verificando instalação..."
google-chrome --version

SSH_EOF
```

### 2️⃣ Instalar ChromeDriver (compatível)
```bash
ssh -i chave.key ubuntu@132.145.57.133 << 'SSH_EOF'

echo "Instalando ChromeDriver..."
CHROME_VERSION=$(google-chrome --version | awk '{print $NF}' | cut -d'.' -f1)
CHROMEDRIVER_URL="https://chromedriver.chromium.org/download"
# ou usar: https://googlechromelabs.github.io/chrome-for-testing/

# Opção fácil: usar apt
sudo apt-get install -y chromium-chromedriver

# Verificar
which chromedriver
chromedriver --version

SSH_EOF
```

### 3️⃣ Testar SOMA após Chrome instalado
```bash
# Rodar manualmente
ssh -i chave.key ubuntu@132.145.57.133 'cd ~/soma-automation/SOMA && .venv/bin/python main.py'

# Ou via PM2
ssh -i chave.key ubuntu@132.145.57.133 'pm2 restart soma-automation'
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
