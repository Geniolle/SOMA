# 🤖 ChromeDriver Auto-Install (Selenium Manager)

**Descoberta:** ChromeDriver pode ser instalado automaticamente pelo Selenium!

---

## 🎯 Como Funciona

O SOMA usa Selenium 4.43.0 que tem **Selenium Manager** integrado.

### Fluxo de Inicialização

1. Python executa `main.py`
2. SOMA cria WebDriver Chrome
3. Selenium Manager detecta Chrome 146.0
4. Selenium Manager auto-download ChromeDriver 146
5. ChromeDriver inicia e conecta ao Chrome
6. SOMA começa a automação

### Locais de Cache

ChromeDriver é baixado para:
```
~/.cache/selenium/            # Selenium Manager cache (Linux)
~/.wdm/                        # webdriver-manager cache
```

---

## ✅ Por que Não Precisa Instalar Manualmente

### Dependências do Projeto
```
selenium==4.43.0      # ← Tem Selenium Manager built-in
webdriver-manager==4.0.2  # ← Fallback alternativo
```

### Selenium Manager (Built-in do Selenium 4)
```python
# Sem precisar configurar nada, Selenium Manager:
from selenium import webdriver

driver = webdriver.Chrome()  # ✅ Auto-download do ChromeDriver!
```

### webdriver-manager (Fallback)
```python
# Se Selenium Manager falhar, webdriver-manager pega:
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
```

---

## 🚀 Como Rodar SOMA Agora

### Opção 1: Direto via SSH (Recomendado)
```bash
ssh -i chave.key ubuntu@132.145.57.133 << 'EOF'
cd ~/soma-automation/SOMA
.venv/bin/python main.py
EOF
```

**O que acontece:**
1. Python inicia
2. Selenium Manager detecta Chrome 146
3. Selenium Manager auto-download ChromeDriver 146
4. ChromeDriver inicia
5. SOMA começa a rodar

### Opção 2: Via PM2 (Melhor para Produção)
```bash
ssh -i chave.key ubuntu@132.145.57.133 << 'EOF'
cd ~/soma-automation/SOMA
pm2 restart soma-automation
pm2 logs soma-automation
EOF
```

### Opção 3: Test Script (Melhor para Verificar)
```bash
ssh -i chave.key ubuntu@132.145.57.133 'cd ~/soma-automation/SOMA && bash test-soma.sh'
```

---

## 📊 Verificação

Depois de rodar, procure por:

### ✅ Sucesso
```
Validacao DOC.SOMA: 9 documentos disponiveis para lancamento.
Processando todas as linhas!
```

### ⏳ Ainda Carregando
```
2026-08-05 22:00:00 | INFO | soma_app.workflows.run_soma | Starting SOMA automation...
2026-08-05 22:00:05 | INFO | soma_app.infra.webdriver_factory | ChromeDriver downloading...
```

### ⚠️ Erro Esperado (primeira vez)
```
SessionNotCreatedException: session not created from chrome not reachable
```
→ Chrome pode estar demorando para iniciar  
→ Rodar novamente: `pm2 restart soma-automation`

---

## 🔍 Técnico: Como Selenium Manager Funciona

### 1. Detecção de Chrome
```
SOMA inicia → Selenium detecta Chrome 146.0
                ↓
        Precisa de ChromeDriver 146
```

### 2. Auto-Download
```
wget https://edgedl.me.chromium.org/chromedriver/146.0.x/chromedriver
  ↓
Salva em ~/.cache/selenium/chromedriver
  ↓
Chdir +x permission
```

### 3. Verificação de Versão
```python
chrome_version = 146.0
chromedriver_version = 146.0
✅ Match! Pode usar
```

### 4. Inicialização
```
webdriver.Chrome(service=Service())
  ↓
Service() procura chromedriver em:
  1. PATH environment
  2. Selenium Manager cache ✅
  3. webdriver-manager cache
  ↓
Encontrado em ~/.cache/selenium/
  ↓
ChromeDriver inicia na porta 9515
```

---

## 📋 Checklist

- [x] Chrome 146.0 instalado no servidor
- [x] Selenium 4.43.0 tem Selenium Manager built-in
- [x] webdriver-manager 4.0.2 disponível como fallback
- [x] SOMA configurado para usar auto-download
- [x] Permissões de venv corretas
- [x] Git sincronizado

**Próximo:** Apenas rodar SOMA e deixar Selenium Manager fazer seu trabalho!

---

## 🎯 Comando Final para Ativar

```bash
# Simples e direto:
ssh -i chave.key ubuntu@132.145.57.133 << 'EOF'
cd ~/soma-automation/SOMA
pm2 restart soma-automation
sleep 3
pm2 logs soma-automation --lines 50 --nostream
EOF
```

**ETA:** 
- Restart PM2: 2 seg
- Download ChromeDriver: 30-60 seg (primeira vez)
- SOMA inicia: 10-20 seg
- **Total:** 1-2 minutos

---

## 📚 Referência

- Selenium Manager: https://www.selenium.dev/selenium-manager/
- webdriver-manager: https://github.com/SergeyPirogov/webdriver_manager
- SOMA WebDriver Factory: src/soma_app/infra/webdriver_factory.py

---

**Conclusão:** Não precisa instalar ChromeDriver manualmente! Selenium faz isso automaticamente. 🚀
