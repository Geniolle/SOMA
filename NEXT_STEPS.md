# ⚡ Próximos Passos - Ativar SOMA em Produção

**Status Atual:** ChromeDriver sendo instalado no servidor  
**ETA:** 2-3 minutos  
**Responsável:** Automático via apt-get  

---

## 📍 Status em Tempo Real

### ✅ Já Concluído
- [x] Chrome 146.0.7680.153 instalado no servidor
- [x] Git sincronizado com latest code
- [x] Python venv criado
- [x] requirements.lock.txt instalado
- [x] PM2 configurado (ecosystem.config.js)
- [x] soma-automation registrado no PM2

### ⏳ Em Andamento
- [ ] ChromeDriver sendo instalado (apt-get install chromium-chromedriver)

### ⏹️ Aguardando Conclusão
- [ ] SOMA reiniciar via PM2
- [ ] Verificar logs para erro de Chrome
- [ ] Iniciar automático (se tudo OK)

---

## 🎯 O que fazer quando ChromeDriver terminar

### Opção A: Manual (Recomendado para teste)
```bash
ssh -i chave.key ubuntu@132.145.57.133 << 'EOF'
cd ~/soma-automation/SOMA
bash test-soma.sh
EOF
```

**O que faz:**
- Verifica Chrome ✅
- Verifica ChromeDriver ✅
- Reinicia SOMA ✅
- Mostra logs ✅

### Opção B: Direto via PM2
```bash
ssh -i chave.key ubuntu@132.145.57.133 'pm2 restart soma-automation && pm2 logs soma-automation --lines 50 --nostream'
```

### Opção C: Monitoramento em Tempo Real
```bash
ssh -i chave.key ubuntu@132.145.57.133 'pm2 logs soma-automation'
```

---

## ✅ Verificação de Sucesso

Depois de rodar test-soma.sh, procure por:

### ✅ Sucesso
```
✅ Google Chrome 146.0.7680.153
✅ ChromeDriver version...
✅ soma-automation restarted
Status: online (green)
```

### ❌ Erro Esperado (primeiro startup)
```
[STEP_FAIL] run.init | SessionNotCreatedException
```
→ Significa que está tentando conectar ao Chrome (esperado)  
→ Chrome pode demorar para inicia a primeira vez

### ❌ Erro Crítico
```
session not created from chrome not reachable
```
→ Chrome ainda não foi encontrado  
→ Rodar novamente: `bash test-soma.sh`

---

## 📊 Timeline Esperada

| Passo | Tempo | Status |
|-------|-------|--------|
| 1. ChromeDriver apt-get | 1-2 min | ⏳ Em Andamento |
| 2. ChromeDriver verificação | 30 seg | ⏹️ Aguardando |
| 3. PM2 restart SOMA | 10 seg | ⏹️ Aguardando |
| 4. SOMA inicia Chrome | 5-10 seg | ⏹️ Aguardando |
| 5. SOMA conecta ao site | 30-60 seg | ⏹️ Aguardando |
| **TOTAL** | **3-5 min** | ⏳ |

---

## 🆘 Troubleshooting Rápido

### Se ChromeDriver falhar na instalação
```bash
# Verificar lock apt
ssh -i chave.key ubuntu@132.145.57.133 'sudo lsof /var/lib/apt/lists/lock'

# Liberar lock
ssh -i chave.key ubuntu@132.145.57.133 'sudo rm /var/lib/apt/lists/lock'

# Tentar novamente
ssh -i chave.key ubuntu@132.145.57.133 'sudo apt-get install -y chromium-chromedriver'
```

### Se SOMA não inicia após ChromeDriver
```bash
# Ver logs completos
ssh -i chave.key ubuntu@132.145.57.133 'pm2 logs soma-automation --lines 100 --nostream'

# Restart manual
ssh -i chave.key ubuntu@132.145.57.133 'pm2 stop soma-automation && pm2 start ecosystem.config.js'

# Health check
ssh -i chave.key ubuntu@132.145.57.133 'bash ~/soma-automation/SOMA/control-soma.sh health'
```

### Se Chrome não encontrado
```bash
# Verificar Chrome instalado
ssh -i chave.key ubuntu@132.145.57.133 'google-chrome --version'

# Se não encontrar, rodar install-chrome.sh novamente
ssh -i chave.key ubuntu@132.145.57.133 'bash ~/soma-automation/SOMA/install-chrome.sh'
```

---

## 🎯 Comandos de Monitoramento Pós-Startup

Depois que SOMA estiver rodando:

### Dashboard em Tempo Real
```bash
ssh -i chave.key ubuntu@132.145.57.133 'pm2 monit'
```

### Logs Contínuos
```bash
ssh -i chave.key ubuntu@132.145.57.133 'pm2 logs soma-automation'
```

### Últimas 50 Linhas
```bash
ssh -i chave.key ubuntu@132.145.57.133 'pm2 logs soma-automation --lines 50 --nostream'
```

### Health Check Completo
```bash
ssh -i chave.key ubuntu@132.145.57.133 'bash ~/soma-automation/SOMA/control-soma.sh health'
```

### Status Geral
```bash
ssh -i chave.key ubuntu@132.145.57.133 'pm2 status'
```

---

## 📞 Contato de Suporte

Se tudo falhar:

1. Verificar arquivo de log: `~/soma-automation/SOMA/logs/soma-error.log`
2. Revisar SERVER_ANALYSIS.md para troubleshooting
3. Revisar INSTALL_GUIDE.md para pré-requisitos
4. Revisar DEPLOYMENT_COMPLETE.md para contexto

---

## 🎉 Resultado Esperado (Sucesso)

Depois que SOMA estiver rodando:

```
[soma-automation] 2026-08-05 22:00:00 | INFO | soma_app.workflows.run_soma | Starting SOMA automation...
[soma-automation] 2026-08-05 22:00:05 | INFO | soma_app.automation.pages.login_page | Initiating login...
[soma-automation] 2026-08-05 22:00:10 | INFO | soma_app.automation.pages.login_page | Login successful
[soma-automation] 2026-08-05 22:00:15 | INFO | soma_app.automation.actions | Processing data...
[soma-automation] 2026-08-05 22:00:20 | INFO | soma_app.automation.pages.entradas_saidas_page | Inserting payment data...
```

---

**Próxima ação:** Aguardar notificação de conclusão do ChromeDriver (1-2 minutos)  
**Então:** Rodar `bash test-soma.sh` via SSH  
**Resultado:** SOMA automático em produção 🚀
