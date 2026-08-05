# ✅ XPaths Corretos - Modal "Inserir Pagamento"

**Data:** 2026-08-05  
**Status:** 🟢 ATUALIZADO COM XPATHS CORRETOS

---

## 📋 XPaths Extraídos Manualmente

### 1️⃣ Botão: INSERIR PAGAMENTO (BTN_INSERIR_BAIXA)
```xpath
/html/body/div[2]/div/div[2]/div/div/form/div[29]/div/div[2]/a
```
- **Tag:** `<a>` (link)
- **Classe:** `btn btn-info btn-block bnt_inserir`
- **Ação:** Abre o modal de pagamento

---

### 2️⃣ Campo: DATA PAGAMENTO (DATA_BAIXA)
```xpath
/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input
```
- **Tag:** `<input type="text">`
- **Classe:** `form-control date datepicker`
- **Formato:** `dd/mm/yyyy`
- **Preenchimento:** `row.data_mov`

**Observação:** O erro anterior era `div[4]`, o correto é `div[5]` ✅

---

### 3️⃣ Campo: FORMA DE PAGAMENTO
```xpath
/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select
```
- **Tag:** `<select>`
- **Atributo:** `name="forma_pagamento"`
- **Tipo:** Dropdown com opções:
  - DINHEIRO
  - DEPÓSITO
  - CHEQUE
  - TRANSFERÊNCIA BANCÁRIA
  - PIX
  - MAQUINETA

---

### 4️⃣ Botão: SALVAR PAGAMENTO (BTN_SALVAR_BAIXA)
```xpath
/html/body/div[2]/div/div[5]/div/div/form/div[3]/button
```
- **Tag:** `<button type="submit">`
- **Texto:** "Salvar Baixa"
- **Classe:** `btn btn-info`
- **Ação:** Salva o pagamento e fecha o modal

---

## 📊 Estrutura do Modal

```
Modal #inserir
├─ Form (id="form_pagamentos")
│  ├─ Modal Header
│  │  └─ Título: "Adicione um pagamento para esta conta"
│  │
│  ├─ Modal Body
│  │  ├─ div[2]: Seção de Data
│  │  │  ├─ DATA PAGAMENTO (input)
│  │  │  └─ Aplicar Desconto (checkbox)
│  │  │
│  │  ├─ div[3]: Seção de Forma de Pagamento
│  │  │  ├─ FORMA DE PAGAMENTO (select)
│  │  │  ├─ N° CHEQUE (input - escondido)
│  │  │  └─ N° DOCUMENTO (input - escondido)
│  │  │
│  │  └─ Mais campos...
│  │
│  └─ Modal Footer
│     ├─ BTN_SALVAR_BAIXA (button)
│     └─ Inputs ocultos
```

---

## 🔄 Fluxo Completo Atualizado

```python
def _do_baixa(self, row: ContaOrdemRow):
    # 1. Clica em "Inserir Pagamento"
    self.a.click_js(BTN_INSERIR_BAIXA)  # /html/body/div[2]/div/div[2]/div/div/form/div[29]/div/div[2]/a
    time.sleep(1)
    
    # 2. Preenche DATA PAGAMENTO (data da baixa)
    self.a.type(DATA_BAIXA, row.data_mov)
    # Xpath: /html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input
    time.sleep(0.5)
    
    # 3. Seleciona FORMA DE PAGAMENTO (se necessário)
    # self.a.select_by_text(FORMA_PAGAMENTO_MODAL, "DINHEIRO")
    # Xpath: /html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select
    
    # 4. Clica em popup (se aparecer)
    try:
        loc = self.a.wait_any_present(POPUP_CLICK_CANDIDATES, timeout_seconds=3)
        self.a.click_js(loc)
    except Exception:
        pass
    
    # 5. Clica em "Salvar Baixa"
    self.a.click_js(BTN_SALVAR_BAIXA)
    # Xpath: /html/body/div[2]/div/div[5]/div/div/form/div[3]/button
    time.sleep(2)
```

---

## 📝 Diferenças Encontradas

| Campo | Antes (Errado ❌) | Depois (Correto ✅) | Diferença |
|-------|-------------------|-------------------|-----------|
| Modal Div | `div[4]` | `div[5]` | Incremento de 1 |
| Data Input | `/div/div/input` | `/div[1]/div/input` | Nível extra `[1]` |

---

## ✅ Arquivos Atualizados

### 1. `src/soma_app/automation/pages/entradas_saidas_page.py`
```python
DATA_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input")
FORMA_PAGAMENTO_MODAL = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select")
POPUP_CLICK_CANDIDATES = []
BTN_SALVAR_BAIXA = (By.XPATH, "/html/body/div[2]/div/div[5]/div/div/form/div[3]/button")
```

### 2. `src/soma_app/config/locators.json`
```json
"DATA_BAIXA": "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[1]/div[1]/div/input",
"FORMA_PAGAMENTO_MODAL": "/html/body/div[2]/div/div[5]/div/div/form/div[2]/div/div/div[2]/div[1]/div/select",
"POPUP_CLICK_CANDIDATES": [...],
"BTN_SALVAR_BAIXA": "/html/body/div[2]/div/div[5]/div/div/form/div[3]/button"
```

---

## 🚀 Resultado Esperado

✅ Campo DATA PAGAMENTO agora será encontrado (antes dava timeout)  
✅ Sistema conseguirá preencher a data  
✅ Conseguirá clicar em "Salvar Baixa"  
✅ Modal fechará e fluxo continuará  

---

## 📌 Próximas Etapas

1. Testar com `python main.py`
2. Monitorar se DATA_BAIXA é preenchido com sucesso
3. Se houver mais campos para preencher no modal, extrair seus xpaths
4. Confirmar que "Salvar Baixa" é clicado com sucesso

---

**Status:** 🟢 XPATHS EXTRAÍDOS E ATUALIZADOS COM SUCESSO!
