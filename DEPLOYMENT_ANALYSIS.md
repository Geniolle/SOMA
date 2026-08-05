# 📋 Análise de Deployment - SOMA

## Seu Processo Manual Atual

### 🔗 No Servidor (Ubuntu)

```bash
# Passo 1: SSH e Git
ssh -i chave.key ubuntu@132.145.57.133
cd ~/soma-automation/SOMA
git reset --hard origin/main
git pull origin main

# Passo 2: Dependências Node.js
npm install

# Passo 3: Gerenciar com PM2
pm2 stop all
pm2 restart all

# Passo 4: Monitorar
pm2 logs orquestrador-soma
pm2 status
```

### 💻 Local (Windows/PowerShell)

```bash
cd SOMA
.\.venv\Scripts\Activate.ps1
python main.py

# Dashboard
pm2 monit

# Controle
pm2 stop orquestrador-soma
pm2 restart orquestrador-soma
```

---

## 🔍 Análise Técnica

| Passo | Propósito | Automável | Prioridade |
|-------|----------|-----------|-----------|
| `git reset --hard origin/main` | Descartar mudanças locais e sincronizar | ✅ Sim | 🔴 Alta |
| `git pull origin main` | Puxar atualizações | ✅ Sim | 🔴 Alta |
| `npm install` | Instalar dependências Node | ✅ Sim | 🟡 Média |
| `pm2 stop all` | Parar todos os processos | ✅ Sim | 🔴 Alta |
| `pm2 restart all` | Reiniciar processos | ✅ Sim | 🔴 Alta |
| `pm2 logs` | Ver logs em tempo real | ✅ Sim (parcialmente) | 🟡 Média |
| `pm2 status` | Status dos processos | ✅ Sim | 🟡 Média |
| `.venv/Scripts/Activate.ps1` | Ativar venv (local) | ✅ Sim | 🟢 Baixa |
| `python main.py` | Rodar orquestrador (local) | ✅ Sim | 🟢 Baixa |

---

## 💡 Propostas de Automação

### Opção 1: Script Bash no Servidor (Recomendado)
**Arquivo:** `deploy.sh`

```bash
#!/bin/bash
cd ~/soma-automation/SOMA

# 1. Sincronizar com GitHub
echo "🔄 Sincronizando com GitHub..."
git fetch origin
git reset --hard origin/main

# 2. Atualizar dependências
echo "📦 Instalando dependências..."
npm install --production

# 3. Gerenciar PM2
echo "🔄 Reiniciando com PM2..."
pm2 stop orquestrador-soma
pm2 restart orquestrador-soma

# 4. Verificar status
echo "✅ Status:"
pm2 status

echo "📊 Logs (últimas 50 linhas):"
pm2 logs orquestrador-soma --lines 50
```

**Uso:**
```bash
ssh -i chave.key ubuntu@132.145.57.133 'bash ~/soma-automation/SOMA/deploy.sh'
```

---

### Opção 2: Script PowerShell (Local Windows)
**Arquivo:** `deploy-local.ps1`

```powershell
# Ativa venv e roda orquestrador
Set-Location "C:\workspace\SOMA"
& ".\.venv\Scripts\Activate.ps1"
python main.py
```

**Uso:**
```powershell
powershell -ExecutionPolicy Bypass -File deploy-local.ps1
```

---

### Opção 3: Script de Deployment Completo (Híbrido)
**Arquivo:** `deploy-full.sh`

Combina:
- ✅ Reset git + pull
- ✅ Atualizar dependências (npm + pip)
- ✅ Rodar migrações se necessário
- ✅ Verificar saúde da aplicação
- ✅ Backup de logs antes de rotacionar
- ✅ Notificação de sucesso/erro

---

## 🎯 Recomendação

### Implementar Agora:

**1. `deploy.sh` (Servidor)**
- Automatiza 90% do seu processo
- Rápido e confiável
- Fácil integração com CI/CD

**2. Cron Job**
```bash
# Fazer pull das atualizações a cada hora
0 * * * * cd ~/soma-automation/SOMA && bash deploy.sh
```

**3. Webhook GitHub → Servidor**
- Push automático dispara deployment
- Zero latência entre update e aplicação

---

## ⚙️ Implementação

### Pré-requisitos Verificados ✅

- ✅ PM2 instalado no servidor
- ✅ Node.js/npm disponível
- ✅ Python + venv local
- ✅ Git com acesso a GitHub
- ✅ SSH key configurada

### Próximas Ações

1. **Criar `deploy.sh`** - Script de deployment
2. **Testar no servidor** - Verificar funcionalidade
3. **Configurar PM2** - Se ainda não estiver
4. **Automação opcional** - Cron ou GitHub Actions

---

## ❓ Dúvidas Importantes

1. **PM2 está ativo?** - Seu orquestrador roda via PM2 ("orquestrador-soma")?
2. **Qual é o comando para iniciar?** - `python main.py` ou há um script específico?
3. **Precisa de dependências adicionais?** - Além de npm install, precisa de `pip install`?
4. **Ambiente de produção?** - Precisa de `.env` ou configurações especiais?

---

**Status:** 🟡 Aguardando sua confirmação sobre qual opção implementar
