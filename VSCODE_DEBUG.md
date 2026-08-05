# 🐛 Debug no VSCode: Guia Completo

**Arquivo de Configuração:** `.vscode/launch.json`  
**Script de Debug:** `debug_locators.py`

---

## 🚀 Modo 1: Debug Simples (Recommended)

### 1. Abrir VSCode
```bash
code .
```

### 2. Ir para Debug (Ctrl+Shift+D)
- Clique no ícone de Debug na sidebar esquerda
- Ou pressione `Ctrl+Shift+D`

### 3. Selecionar Configuração
- Dropdown no topo: **"SOMA Debug"**
- Ou **"SOMA Debug (with breakpoints)"** para mais detalhe

### 4. Iniciar Debug
- Clique no botão Play (▶️) ou pressione `F5`

### 5. Terminal Integrado Abre
- Script executa no terminal integrado
- Pausas interativas funcionam aqui também

---

## 🎯 Modo 2: Debug com Breakpoints

### 1. Adicionar Breakpoint
- Clique na linha de código onde quer pausar
- Aparece um ponto vermelho

### 2. Iniciar Debug
- F5 ou clique em Play

### 3. Executar até o Breakpoint
- Script executa até chegar no breakpoint
- Para automaticamente

### 4. Inspecionar Variáveis
- **Variables (esquerda):** Ver variáveis locais
- **Watch:** Adicionar expressões para monitorar
- **Debug Console (Ctrl+Shift+Y):** Executar código Python

### 5. Controles
- **Continue:** F5 (continua até próximo breakpoint)
- **Step Over:** F10 (próxima linha)
- **Step Into:** F11 (entra em função)
- **Step Out:** Shift+F11 (sai de função)

---

## 🔍 Modo 3: Debug de Locators (Recomendado Primeiro)

Usar o script específico para verificar se os locators estão carregados:

### 1. Executar Script de Debug
```bash
python debug_locators.py
```

### 2. Com Debug no VSCode
- Abrir `debug_locators.py`
- F5 para executar com debug
- Ver output no terminal integrado

### 3. Verificar Output
Deve mostrar:
```
✅ BTN_INSERIR_BAIXA_CANDIDATES encontrado no JSON!
📋 Quantidade de candidates: 5
   1. //table//button[contains(@title,'Inserir')]
   2. //table//tbody//tr[1]//td[6]//button
   ...
```

Se mostrar:
```
❌ BTN_INSERIR_BAIXA_CANDIDATES NÃO encontrado no JSON!
```

Então temos um problema na configuração do JSON.

---

## 📊 Debug Console Útil

Quando parado em um breakpoint, use Debug Console (Ctrl+Shift+Y):

```python
# Ver variável
page.BTN_INSERIR_BAIXA_CANDIDATES

# Ver tipo
type(page.BTN_INSERIR_BAIXA_CANDIDATES)

# Ver tamanho
len(page.BTN_INSERIR_BAIXA_CANDIDATES)

# Iterar
for candidate in page.BTN_INSERIR_BAIXA_CANDIDATES:
    print(candidate)
```

---

## 🎯 Problema: BTN_INSERIR_BAIXA_CANDIDATES Vazio

### Sintomas
```
[PRE-BAIXA] Nenhum candidate de BTN_INSERIR_BAIXA encontrado. Candidates: []
```

### Debug Steps

1. **Verificar JSON**
   ```bash
   python debug_locators.py
   ```
   - Procurar por "BTN_INSERIR_BAIXA_CANDIDATES"
   - Verificar se está no JSON

2. **Adicionar Breakpoint em `apply_locator_overrides`**
   - Arquivo: `src/soma_app/config/locators.py`
   - Função: `apply_locator_overrides()`
   - Linha: onde percorre os locators

3. **Executar com Debug**
   - F5 para iniciar
   - Ver variáveis no lado esquerdo
   - Verificar se JSON foi carregado

4. **Inspecionar no Console**
   ```python
   cfg  # Ver config carregada
   "BTN_INSERIR_BAIXA_CANDIDATES" in cfg  # True/False
   page.BTN_INSERIR_BAIXA_CANDIDATES  # Ver valor
   ```

---

## 🔧 Exemplos de Breakpoints Úteis

### 1. Debug de Locator Loading
- Arquivo: `src/soma_app/config/locators.py`
- Função: `load_page_locator_config()`
- Ver qual JSON foi carregado

### 2. Debug de Aplicação de Overrides
- Arquivo: `src/soma_app/config/locators.py`
- Função: `apply_locator_overrides()`
- Ver quais locators foram aplicados

### 3. Debug de Ação
- Arquivo: `src/soma_app/automation/actions.py`
- Função: `wait_any_present()`
- Ver quais candidates estão sendo procurados

### 4. Debug de Preenchimento
- Arquivo: `src/soma_app/automation/pages/entradas_saidas_page.py`
- Função: `_fill_common()`
- Ver contexto de debug ativado

---

## 💻 Teclado Útil

| Tecla | Ação |
|-------|------|
| F5 | Iniciar/Continuar debug |
| F6 | Pausar |
| F9 | Toggle breakpoint |
| F10 | Step over |
| F11 | Step into |
| Shift+F11 | Step out |
| Ctrl+Shift+D | Abrir Debug |
| Ctrl+Shift+Y | Debug Console |
| Ctrl+` | Terminal Integrado |

---

## 📋 Checklist de Debug

- [ ] `.vscode/launch.json` existe
- [ ] VSCode consegue encontrar Python (Ctrl+Shift+P > Python: Select Interpreter)
- [ ] Executar `debug_locators.py` mostra candidates
- [ ] Adicionar breakpoint em `_do_baixa()`
- [ ] F5 para iniciar debug
- [ ] Parar no breakpoint
- [ ] Inspecionar `self.BTN_INSERIR_BAIXA_CANDIDATES` no Console
- [ ] Ver se está vazio ou tem 5 candidates

---

## 🆘 Problemas Comuns

### "Python not found"
```
Ctrl+Shift+P > Python: Select Interpreter
Escolher: .venv/Scripts/python.exe
```

### "launch.json not found"
```
Ctrl+Shift+D > Create a launch.json file
Escolher: Python
```

### Breakpoint não para
```
Verificar: Debug Console mostra "Breakpoint set"
Ter certeza que é arquivo certo
```

### Variáveis não aparecem
```
Pode ser variáveis locais perdidas
Usar Debug Console para inspecionar
```

---

## 🎓 Tutorial Completo

### Passo 1: Setup
```bash
code .  # Abrir VSCode
```

### Passo 2: Verificar Locators
```bash
Ctrl+` (abrir terminal)
python debug_locators.py
```

### Passo 3: Debug Completo
- Ctrl+Shift+D
- Selecionar "SOMA Debug"
- F5
- Acompanhar logs

### Passo 4: Adicionar Breakpoint (se necessário)
- Abrir arquivo
- Clique na linha
- Ponto vermelho aparece
- F5 novamente
- Para no breakpoint

---

**✅ Pronto para debugar!**

Use F5 para iniciar e Ctrl+Shift+Y para inspecionar variáveis.
