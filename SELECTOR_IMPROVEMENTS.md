# 🔧 Melhorias de Seletores: BTN_INSERIR_BAIXA

**Data:** 2026-08-05  
**Status:** ✅ Implementado

---

## 📋 Problema Original

**Seletor Anterior:**
```xpath
/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button
```

**Problemas:**
- ❌ Muito específico - procura **apenas a primeira linha** da tabela
- ❌ Falha se não houver primeira linha ou se estiver vazia
- ❌ Não reutilizável - muda com a estrutura do DOM
- ❌ Sem fallback - apenas um xpath

**Erro Observado:**
```
TimeoutException: Message:
File "entradas_saidas_page.py", line 767, in _do_baixa
    self.a.click_js(self.BTN_INSERIR_BAIXA)
    
[TIMEOUT] TIMEOUT em click_js | found=0
```

---

## ✅ Solução Implementada

### 1. **Múltiplos Candidates (Fallback Robusto)**

**Nova Abordagem:** `BTN_INSERIR_BAIXA_CANDIDATES` em `locators.json`:

```json
"BTN_INSERIR_BAIXA_CANDIDATES": [
  "//table//button[contains(@title,'Inserir') or contains(.,'Inserir')]",
  "//table//tbody//tr[1]//td[6]//button",
  "//div[@class='table-responsive']//button[contains(@title,'Inserir')]",
  "/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button",
  "//button[@type='button' and (contains(.,'Inserir') or contains(@title,'Inserir'))]"
]
```

### Cada Candidate Explica-se:

| # | XPath | Benefício |
|---|-------|----------|
| 1 | `//table//button[contains(@title,'Inserir') or contains(.,'Inserir')]` | Procura por texto/atributo - **mais semântico** |
| 2 | `//table//tbody//tr[1]//td[6]//button` | Similar ao original mas mais flexível |
| 3 | `//div[@class='table-responsive']//button[contains(@title,'Inserir')]` | Procura por class container |
| 4 | `/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button` | Original (último recurso) |
| 5 | `//button[@type='button' and ...]` | Genérico por atributo |

### 2. **Uso com `wait_any_present()`**

**Código Novo em `entradas_saidas_page.py`:**

```python
def _do_baixa(self, row: ContaOrdemRow) -> None:
    with step(log, "entradas_saidas.baixa", ...):
        self._dismiss_overlays()
        
        # Tenta múltiplos candidates em sequência
        try:
            btn_low = self.a.wait_any_present(
                self.BTN_INSERIR_BAIXA_CANDIDATES, 
                timeout_seconds=30
            )
            log.info("[PRE-BAIXA] BTN_INSERIR_BAIXA encontrado usando: %s", btn_low)
            self.a.click_js(btn_low)
        except TimeoutException:
            log.error("[PRE-BAIXA] Nenhum candidate encontrado")
            raise
```

**Fluxo:**
```
1. Tenta candidate 1 (semântico)
   ↓
2. Se falhar, tenta candidate 2
   ↓
3. Se falhar, tenta candidate 3
   ↓
4. Se falhar, tenta candidate 4 (original)
   ↓
5. Se falhar, tenta candidate 5 (genérico)
   ↓
6. Se todos falharem → TimeoutException com diagnóstico completo
```

---

## 🎯 Benefícios da Solução

| Benefício | Impacto |
|-----------|--------|
| **Semântico** | Procura por texto "Inserir", não por posição HTML |
| **Robusto** | Múltiplos candidates para fallback |
| **Flexível** | Funciona mesmo se DOM mudar um pouco |
| **Debugável** | Log mostra qual candidate funcionou |
| **Reutilizável** | Padrão pode ser aplicado a outros seletores |

---

## 📊 Comparação

### ❌ ANTES (Problema)
```
XPath Único → Falha → Timeout
```

### ✅ DEPOIS (Solução)
```
Candidate 1 (Semântico) → Sucesso ✓
ou
Candidate 1 (Falha) → Candidate 2 → Sucesso ✓
ou
Candidate 1-4 (Falha) → Candidate 5 (Genérico) → Sucesso ✓
ou
Todos (Falha) → TimeoutException com diagnostics ✓
```

---

## 🧪 Como Testar

### 1. **Com Debug Ativo**

```bash
# Editar deploy/.env
DEBUG_SELECTOR_INTERACTIVE=true

# Executar
python main.py
```

**Esperado no log:**
```
[PRE-BAIXA] Tentando clicar em BTN_INSERIR_BAIXA | candidates=5 | url=...
[PRE-BAIXA] BTN_INSERIR_BAIXA encontrado usando: (By.XPATH, '//table//button...')
[SELECTOR] action=click_js | method=xpath | selector=//table//button[contains...
→ Pressione ENTER para continuar...
```

### 2. **Simular Falha**

Se mesmo assim falhar:

1. Verificar arquivo HTML diagnóstico:
   ```bash
   type artifacts/diagnostics/timeout_*.html | findstr "Inserir"
   ```

2. Abrir screenshot:
   ```bash
   start artifacts/screenshots/timeout_*.png
   ```

3. No DevTools (F12), testar:
   ```javascript
   $x("//table//button[contains(.,'Inserir')]")[0]  // Deve aparecer
   ```

---

## 📝 Exemplo de Log Esperado

### Sucesso
```
2026-08-05 16:20:45 | [PRE-BAIXA] Tentando clicar em BTN_INSERIR_BAIXA | candidates=5 | url=https://verbodavida.info/IVV/
2026-08-05 16:20:46 | [PRE-BAIXA] BTN_INSERIR_BAIXA encontrado usando: (By.XPATH, '//table//button[contains(@title,"Inserir")]')
2026-08-05 16:20:47 | [SELECTOR] action=click_js | method=xpath | selector=//table//button[contains(@title,"Inserir")]

✓ action=click_js | method=xpath | selector=//table//button[contains(@title,"Inserir")]
→ Pressione ENTER para continuar...
```

### Falha (com Diagnostics)
```
2026-08-05 16:20:45 | [PRE-BAIXA] Tentando clicar em BTN_INSERIR_BAIXA | candidates=5 | url=...
2026-08-05 16:20:75 | [PRE-BAIXA] Nenhum candidate encontrado. Candidates: [...]
2026-08-05 16:20:75 | [TIMEOUT] TIMEOUT em click_js | method=xpath | selector=... | found=0 | screenshot=... | html=...
```

**Arquivos Gerados:**
- `artifacts/screenshots/timeout_click_js_xpath.png`
- `artifacts/diagnostics/timeout_click_js_xpath.html`

---

## 🔍 Próximas Melhorias

### Candidates Similares para Outros Seletores

Aplicar o mesmo padrão a:
- `BTN_INSERIR_PAGAMENTO_SAIDA` - adicionar candidates
- `BTN_SALVAR_PAGAMENTO_MODAL` - adicionar candidates
- `BTN_SALVAR_BAIXA` - adicionar candidates
- `DATA_BAIXA` - adicionar candidates

---

## 🎓 Lições Aprendidas

1. **Xpaths semânticos** (por texto/atributo) são mais robustos
2. **Múltiplos candidates** aumentam confiabilidade
3. **Logs detalhados** facilitam debug quando falha
4. **Diagnostics automáticos** (screenshots + HTML) economizam tempo

---

## ✅ Arquivo Alterado

**`src/soma_app/config/locators.json`**
- ❌ Removido: `"BTN_INSERIR_BAIXA": "..."`
- ✅ Adicionado: `"BTN_INSERIR_BAIXA_CANDIDATES": [...]` com 5 candidates

**`src/soma_app/automation/pages/entradas_saidas_page.py`**
- ✅ Adicionado: `BTN_INSERIR_BAIXA_CANDIDATES = []`
- ✅ Modificado: `_do_baixa()` para usar `wait_any_present()` com candidates
- ✅ Adicionado: Logs detalhados antes e durante a busca

---

## 🚀 Status

✅ **Implementado e Pronto para Testar**

Próxima Execução:
```bash
DEBUG_SELECTOR_INTERACTIVE=true python main.py
```

Deve encontrar `BTN_INSERIR_BAIXA` com um dos 5 candidates e continuar a execução.

