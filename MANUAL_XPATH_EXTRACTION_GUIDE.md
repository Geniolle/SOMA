# 📋 Guia de Extração Manual de XPaths - Modal "Inserir Pagamento"

## 🎯 Objetivo
Extrair os xpaths corretos de todos os campos e botões do modal que aparece após clicar em "Inserir Pagamento".

---

## 📍 Passo 1: Navegar até o Modal

1. **Abra o navegador** em modo não-headless (já está configurado: `HEADLESS=false`)
2. **Inicie o processo de automação:**
   ```bash
   cd C:\workspace\SOMA
   python main.py
   ```

3. **O sistema vai:**
   - Login automático
   - Navegar para Entradas/Saídas
   - Abrir formulário de nova saída
   - Preencher os campos iniciais
   - Clicar em "Inserir Pagamento"
   - Parar aqui para você inspecionar

---

## 🔍 Passo 2: Inspecionar Elementos no Chrome DevTools

### Como abrir DevTools:
- Pressione **F12** (ou Ctrl+Shift+I)
- Aparecerá um painel no rodapé

### Elementos a Inspecionar:

#### **1. Campo: DATA_BAIXA (Data da Baixa)**
**Localização visual:** Campo de entrada de data no topo do modal

**Como encontrar:**
- Procure por um `<input type="text"` com placeholder ou label **"DATA DA BAIXA"**
- Pode estar com classe `datepicker` ou `form-control`

**O que copiar:**
```
Clique com botão direito no campo → Inspect ou Inspect Element
Copie o XPath:
  - Clique direito no elemento → Copy → Copy full XPath
  OU
  - Clique direito no elemento → Copy → Copy XPath
```

**Exemplo do que você verá:**
```
/html/body/div[2]/div/div[4]/div/div/form/div[2]/div/div/div[1]/div/div/input
```

**Ou crie um xpath mais robusto:**
```xpath
//input[@name='data_baixa']
//input[contains(@class, 'datepicker')]
//input[@data-date-format='dd/mm/yyyy']
//input[@placeholder='dd/mm/yyyy']
```

---

#### **2. Botão: BTN_SALVAR_BAIXA (Salvar Baixa)**
**Localização visual:** Botão azul "Salvar Baixa" no rodapé do modal

**Como encontrar:**
- Procure por um `<button>` ou `<input type="submit">` com texto **"Salvar Baixa"**

**Xpaths possíveis:**
```xpath
//button[contains(text(), 'Salvar Baixa')]
//button[@class='btn btn-info']
//button[contains(.,'Salvar')]
//form[@id='form_baixas']//button[@type='submit']
```

---

#### **3. Elementos Pop-up (POPUP_CLICK_CANDIDATES)**
**Localização visual:** Pode aparecer após preencher DATA_BAIXA

**Como encontrar:**
- Presencha o campo DATA_BAIXA
- Veja se aparece alguma notificação/popup
- Procure por `<div>` com classes como:
  - `modal-footer`
  - `alert`
  - `notification`
  - `dialog`

**Xpaths possíveis:**
```xpath
//div[@class='modal-footer']
//div[contains(@class, 'alert')]
//div[contains(@class, 'notification')]
//button[@class='btn']
```

---

## 📸 Passo 3: Documentar Tudo

Crie uma lista com cada elemento encontrado no formato:

```markdown
### ELEMENTO: [Nome do campo/botão]

**Localização visual:** [Descrição de onde vê na tela]

**XPath 1 (Mais específico):**
/html/body/...

**XPath 2 (Semântico):**
//elemento[contains(@atributo, 'valor')]

**XPath 3 (Por class/id):**
//elemento[@id='id_especifico']

**XPath 4 (Fallback):**
//elemento[contains(text(), 'Texto')]

**Candidatos sugeridos:**
1. [xpath mais confiável]
2. [xpath alternativo 1]
3. [xpath alternativo 2]
4. [xpath genérico fallback]
```

---

## 🎬 Passo 4: Testar os XPaths no Console

No console do Chrome DevTools (aba "Console"):

```javascript
// Teste um xpath (exemplo DATA_BAIXA)
$x("//input[@name='data_baixa']")  // Pressione Enter

// Se encontrar, vai retornar: [input.form-control.datepicker]
// Se não encontrar, retorna: []
```

**Repita para cada elemento:**
- DATA_BAIXA
- BTN_SALVAR_BAIXA
- POPUP_CLICK_CANDIDATES (todos os candidatos)

---

## 📝 Exemplo Completo de Documentação

```markdown
## CAMPO: DATA_BAIXA

**Localização visual:** Campo de entrada acima do botão "Salvar Baixa"

**Estrutura HTML:**
```html
<div class="modal-body">
  <div class="form-group">
    <label>DATA DA BAIXA:</label>
    <div class="col-md-8">
      <input type="text" 
             name="data_baixa" 
             class="form-control date datepicker" 
             data-date-format="dd/mm/yyyy"
             placeholder="dd/mm/yyyy"
             value="05/08/2026">
    </div>
  </div>
</div>
```

**Candidatos (em ordem de preferência):**
1. `//input[@name='data_baixa']`
2. `//input[contains(@class, 'datepicker') and contains(@name, 'data')]`
3. `//input[@placeholder='dd/mm/yyyy']`
4. `//div[@class='modal-body']//input[@type='text'][1]`


## BOTÃO: BTN_SALVAR_BAIXA

**Localização visual:** Botão azul "Salvar Baixa" no rodapé

**Estrutura HTML:**
```html
<div class="modal-footer">
  <input type="hidden" name="add" value="1">
  <button type="submit" class="btn btn-info">Salvar Baixa</button>
</div>
```

**Candidatos (em ordem de preferência):**
1. `//button[contains(text(), 'Salvar Baixa')]`
2. `//button[@class='btn btn-info' and contains(., 'Salvar')]`
3. `//form[@id='form_baixas']//button[@type='submit']`
4. `//div[@class='modal-footer']//button`
```

---

## 🔄 Passo 5: Atualizar o Código

Após documentar todos os xpaths, atualize os arquivos:

### **Arquivo 1: `src/soma_app/automation/pages/entradas_saidas_page.py`**

```python
# Linha ~104
DATA_BAIXA = (By.XPATH, "//input[@name='data_baixa']")

# Linha ~105
POPUP_CLICK_CANDIDATES = [
    "//button[@class='btn btn-default']",
    "//div[contains(@class, 'alert')]",
]

# Linha ~106
BTN_SALVAR_BAIXA = (By.XPATH, "//button[contains(text(), 'Salvar Baixa')]")
```

### **Arquivo 2: `src/soma_app/config/locators.json`**

```json
"entradas_saidas": {
  "DATA_BAIXA": "//input[@name='data_baixa']",
  "POPUP_CLICK_CANDIDATES": [
    "//button[@class='btn btn-default']",
    "//div[contains(@class, 'alert')]"
  ],
  "BTN_SALVAR_BAIXA": "//button[contains(text(), 'Salvar Baixa')]"
}
```

---

## ✅ Checklist Final

- [ ] Abri o navegador com HEADLESS=false
- [ ] Executei python main.py
- [ ] Sistema parou após clicar em "Inserir Pagamento"
- [ ] Abri DevTools (F12)
- [ ] Encontrei e testei DATA_BAIXA
- [ ] Encontrei e testei BTN_SALVAR_BAIXA
- [ ] Encontrei e testei POPUP_CLICK_CANDIDATES
- [ ] Documentei todos os xpaths em um arquivo
- [ ] Testei cada xpath no console com `$x(...)`
- [ ] Atualizei os arquivos Python e JSON
- [ ] Fiz commit das mudanças

---

## 💡 Dicas Importantes

1. **XPath mais robusto:**
   - Use atributos específicos (name, id, data-*)
   - Evite índices numéricos ([1], [2])
   - Prefira texto ou classes

2. **Teste múltiplos candidatos:**
   - 1º: Mais específico (ex: @name='data_baixa')
   - 2º: Por classe + atributo (ex: contains(@class) and @type)
   - 3º: Por texto (ex: contains(text(), 'Salvar'))
   - 4º: Genérico fallback (ex: //div[@class='modal-footer']//button)

3. **Se não encontrar:**
   - A estrutura HTML pode estar aninhada diferente
   - Procure por tags-pai (div, form, section)
   - Use `contains()` em vez de match exato

---

## 📌 Comando para Depois

Após atualizar os arquivos, execute:

```bash
cd C:\workspace\SOMA
python main.py
```

Se os xpaths estiverem corretos, o sistema vai preencher DATA_BAIXA e clicar em "Salvar Baixa" com sucesso!

---

**Boa sorte com a extração manual! 🎯**
