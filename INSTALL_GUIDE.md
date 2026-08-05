# 📋 Guia de Instalação e Ativação do SOMA no Servidor

**Data:** 2026-08-05  
**Status:** ⏳ Instalação de Chrome em andamento  
**Objetivo:** Ativar SOMA para rodar automático via PM2

---

## 📊 Status Atual

| Componente | Status | Ação |
|-----------|--------|------|
| Git Repo | ✅ Sincronizado | OK |
| Python Venv | ✅ Existe | OK |
| Requirements | ✅ Instalados | OK |
| PM2 Config | ✅ Criado (ecosystem.config.js) | OK |
| Chrome | ⏳ Instalando | Em progresso |
| ChromeDriver | ⏳ Instalando | Em progresso |
| SOMA App | ❌ Aguardando Chrome | Próximo passo |

---

## 🔄 Fluxo de Ativação

### Fase 1: Instalação de Chrome ⏳ EM ANDAMENTO
```bash
# No servidor:
ssh -i chave.key ubuntu@132.145.57.133
cd ~/soma-automation/SOMA
bash install-chrome.sh
```

**O que faz:**
- Adiciona repositório do Chrome
- Instala Google Chrome stable
- Instala ChromeDriver

**Estimado:** 2-3 minutos

---

### Fase 2: Sincronizar e Testar ⏳ PRÓXIMO
```bash
# Sincronizar com GitHub
cd ~/soma-automation/SOMA
git pull origin main

# Opção A: Testar manualmente
.venv/bin/python main.py

# Opção B: Iniciar com PM2
pm2 start ecosystem.config.js
pm2 logs soma-automation
```

---

### Fase 3: Monitorar em Produção ⏳ FINAL
```bash
# Ver status
pm2 status

# Ver logs
pm2 logs soma-automation

# Dashboard
pm2 monit
```

---

## ✅ Checklist de Verificação

Depois que Chrome estiver instalado:

- [ ] `google-chrome --version` retorna versão
- [ ] `chromedriver --version` retorna versão
- [ ] Rodar `python main.py` manualmente
- [ ] Verificar se não há erro de Chrome
- [ ] Se OK: `pm2 start ecosystem.config.js`
- [ ] Verificar `pm2 logs soma-automation`
- [ ] Confirmar processo está online em `pm2 status`

---

## 🆘 Troubleshooting

### Se Chrome não funcionar:

**Erro: "Chrome binary not found"**
```bash
# Verificar localização
which google-chrome
which chromium-browser

# Se não encontrar, reinstalar:
sudo apt-get remove -y google-chrome-stable chromium-chromedriver
bash install-chrome.sh
```

**Erro: "ChromeDriver not compatible"**
```bash
# Obter versão do Chrome
CHROME_VERSION=$(google-chrome --version | awk '{print $NF}' | cut -d'.' -f1)
echo "Chrome version: $CHROME_VERSION"

# Verificar ChromeDriver
chromedriver --version

# Se diferentes, reinstalar ChromeDriver
sudo apt-get remove -y chromium-chromedriver
sudo apt-get install -y chromium-chromedriver
```

**Erro: "Headless mode not working"**
```bash
# Verificar se Chrome tem suporte a headless
google-chrome --headless --dump-dom about:blank

# Se falhar, pode faltar dependências:
sudo apt-get install -y libxi6 libgconf-2-4 libxss1 libappindicator1
```

---

## 📞 Comandos Úteis

```bash
# Status do servidor
pm2 status
pm2 monit

# Logs SOMA
pm2 logs soma-automation
pm2 logs soma-automation --lines 100

# Restart SOMA
pm2 restart soma-automation
pm2 restart soma-automation --force

# Parar SOMA
pm2 stop soma-automation

# Iniciar SOMA
pm2 start ecosystem.config.js

# Remover SOMA de PM2
pm2 delete soma-automation

# Health check
bash ~/soma-automation/SOMA/control-soma.sh health

# Ver logs via SSH
ssh -i chave.key ubuntu@132.145.57.133 'pm2 logs soma-automation --lines 50 --nostream'
```

---

## 📋 Log de Erros Conhecidos

### 1. SessionNotCreatedException
```
selenium.common.exceptions.SessionNotCreatedException: session not created
from chrome not reachable
```
**Causa:** Chrome não instalado  
**Solução:** Rodar `bash install-chrome.sh`

### 2. ChromeDriver Version Mismatch
```
This version of ChromeDriver only supports Chrome version X
```
**Causa:** Versões incompatíveis  
**Solução:** Reinstalar ChromeDriver compatível

### 3. Headless Mode Issues
```
[X] Unable to locate Chrome executable
```
**Causa:** Chrome não encontrado em PATH  
**Solução:** Reinstalar Chrome ou ajustar PATH

---

## 🎯 Resultado Esperado

Após instalação completa:

1. ✅ Chrome funcionando em modo headless
2. ✅ SOMA inicia sem erro de Chrome
3. ✅ PM2 mantém SOMA rodando
4. ✅ Auto-restart se falhar
5. ✅ Logs capturando execução
6. ✅ Monitoring via `pm2 monit`

---

**Próximo passo:** Aguardar conclusão da instalação de Chrome ⏳
