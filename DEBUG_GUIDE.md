# 🔍 Guia de Debug Interativo do SOMA

## Ativação

### Modo Debug Interativo (Pausa após cada ação)

```env
DEBUG_SELECTOR_INTERACTIVE=true
```

**Arquivo:** `deploy/.env`

---

## 📊 O Que é Capturado

### 1. **Logs em Tempo Real (Console)**
```
[SELECTOR] action=click_js | method=xpath | selector=/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button

✓ action=click_js | method=xpath | selector=/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button
→ Pressione ENTER para continuar...
```

### 2. **Arquivo de Log**
- **Localização:** `logs/soma_selectors_YYYYMMDD_HHMMSS.log`
- **Conteúdo:** Timestamp + ação + método + seletor

Exemplo:
```
2026-08-05 16:15:30 | START | modo interativo de seletores ativo
2026-08-05 16:15:31 | action=type | method=name | selector=email | clear=True | value_length=26
2026-08-05 16:15:32 | action=click | method=name | selector=submit
```

### 3. **Diagóstico de Timeout (Automático)**

Quando um seletor falha (TimeoutException):

#### Screenshot
- **Arquivo:** `artifacts/screenshots/timeout_[action]_[method].png`
- **Mostra:** Estado visual da página no momento do timeout

#### HTML Completo
- **Arquivo:** `artifacts/diagnostics/timeout_[action]_[method].html`
- **Mostra:** Código-fonte completo da página para inspeção

#### Log de Erro
```
[TIMEOUT] TIMEOUT em click_js | method=xpath | selector=... | found=0 | screenshot=... | html=... | url=... | title=...
```

---

## 🚀 Como Usar para Debugar Erros

### Cenário 1: Seletor não encontrado (TimeoutException)

**Erro no log:**
```
selenium.common.exceptions.TimeoutException: Message: 
...
File "actions.py", line 226, in click_js
    el = self.wait_present(locator, timeout_seconds=30)
```

**Ações de Debug:**

1. **Ativar modo interativo:**
   ```env
   DEBUG_SELECTOR_INTERACTIVE=true
   ```

2. **Executar novamente** - o script para ANTES de falhar:
   ```
   → Pressione ENTER para continuar...
   ```

3. **Inspencionar o navegador** - a página está visível, você pode:
   - Abrir DevTools (F12)
   - Copiar o xpath do `Console` e testar
   - Verificar se o elemento existe
   - Procurar elementos similares

4. **Arquivos gerados quando falha:**
   - ✓ `artifacts/screenshots/timeout_click_js_xpath.png` - visual
   - ✓ `artifacts/diagnostics/timeout_click_js_xpath.html` - código HTML
   - ✓ `logs/soma_selectors_*.log` - timeline de ações

### Cenário 2: Elemento existe mas não está clicável

**Sintomas:**
- Screenshot mostra o elemento
- Mas `found=0` ou o elemento existe mas não é clicável

**Debug:**

1. No DevTools (F12), teste o xpath:
   ```javascript
   $x("//seu/xpath/aqui")[0]  // Mostra o elemento
   ```

2. Verifique:
   - `element.is_displayed()` - está visível?
   - `element.is_enabled()` - está ativo?
   - Há overlays ocultando? (modais, dropdowns)

3. Verifique no arquivo HTML gerado:
   ```bash
   type artifacts/diagnostics/timeout_*.html | findstr "td\[6\]"
   ```

### Cenário 3: Elemento muda de local (DOM dinâmico)

**Problema:** O xpath é específico de índice: `/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button`

**Solução:**

1. Usar xpath mais robusto no `locators.json`:
   ```json
   "BTN_INSERIR_BAIXA": "//table//button[contains(@title,'Inserir') or contains(.,'Baixa')]"
   ```

2. Usar candidates (múltiplos xpaths):
   ```json
   "BTN_INSERIR_BAIXA_CANDIDATES": [
     "//table//button[contains(@title,'Inserir')]",
     "//button[contains(.,'Baixa')]",
     "/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button"
   ]
   ```

---

## 📁 Estrutura de Arquivos Gerados

```
SOMA/
├── logs/
│   ├── soma_selectors_20260805_161530.log      ← Log interativo
│   ├── soma_dev_20260805_161530.log            ← Log da app
│   └── soma_audit_20260805_161530.log          ← Log de auditoria
│
└── artifacts/
    ├── screenshots/
    │   ├── timeout_click_js_xpath.png          ← Screenshot de erro
    │   └── ...
    │
    └── diagnostics/
        ├── timeout_click_js_xpath.html         ← HTML completo
        └── ...
```

---

## 🎯 Workflow Recomendado

### 1. Primeira Execução (com debug)
```bash
# Ativar debug no .env
DEBUG_SELECTOR_INTERACTIVE=true

# Executar
python main.py

# Quando parar numa ação suspeita:
# - Inspencionar navegador
# - Pressionar ENTER
# - Se falhar, verificar screenshots/html
```

### 2. Corrigir Seletor
```bash
# Editar locators.json
# Testar xpath novo no DevTools
# Atualizar "BTN_INSERIR_BAIXA"
```

### 3. Executar Sem Debug
```bash
# Desativar para velocidade
DEBUG_SELECTOR_INTERACTIVE=false

# Executar
python main.py
```

---

## 🔧 Troubleshooting

### "Pressione ENTER para continuar" não aparece?
- Verificar se `DEBUG_SELECTOR_INTERACTIVE=true` no `.env`
- Verificar se o terminal está ativo (não minimizado)
- Verificar `logs/soma_selectors_*.log`

### Screenshots vazios?
- Pode ser um timeout na captura
- Verificar arquivo HTML em `artifacts/diagnostics/`
- Verificar logs em `logs/soma_dev_*.log`

### HTML muito grande?
- Use busca no browser (Ctrl+F)
- Procure por atributos do elemento (class, id, name)
- Use `grep` para buscar no terminal

---

## 📝 Exemplo Prático

**Erro original:**
```
TimeoutException: Message: 
File "entradas_saidas_page.py", line 767, in _do_baixa
    self.a.click_js(self.BTN_INSERIR_BAIXA)
```

**Com Debug Interativo:**

1. `DEBUG_SELECTOR_INTERACTIVE=true`
2. Script pausa quando chega em BTN_INSERIR_BAIXA
3. Você pressiona F12, copia xpath, e testa
4. Vê que a tabela não foi carregada ainda
5. Aumenta timeout ou adiciona wait para DOM ready
6. Atualiza seletor em `locators.json`
7. Testa novamente até funcionar
8. `DEBUG_SELECTOR_INTERACTIVE=false` para produção

---

## 🎓 Dicas

- **Use `wait_dom_ready()`** antes de interações críticas
- **Use xpaths relativos** ao invés de caminhos absolutos: `//button[contains(.,'Baixa')]` é melhor que `/html/body/div[2]/...`
- **Use `CANDIDATES`** quando houver múltiplas variantes de um elemento
- **Salve screenshots** de estados esperados para comparação
- **Documente** xpaths difíceis com comentários

