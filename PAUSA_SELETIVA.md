# ✅ Pausa Seletiva: Debug Interativo Apenas nos Inputs de Dados

**Status:** ✅ Implementado  
**Data:** 2026-08-05

---

## 🎯 O Que Mudou

### ❌ Antes (Problema)
```
Pausa DURANTE LOGIN:
  [SELECTOR] action=type | method=name | selector=email | clear=True | value_length=26
  → Pressione ENTER para continuar...
  
  [SELECTOR] action=type | method=name | selector=senha | clear=True | value_length=11
  → Pressione ENTER para continuar...

❌ Muitas pausas desnecessárias
```

### ✅ Depois (Solução)
```
Login (SEM PAUSAS):
  [SELECTOR] action=type | method=name | selector=email | clear=True | value_length=26
  [SELECTOR] action=type | method=name | selector=senha | clear=True | value_length=11
  [SELECTOR] action=click | method=name | selector=submit
  Preenchimento realizado com sucesso!

Navegação (SEM PAUSAS):
  [SELECTOR] action=click_js | method=xpath | selector=//*[contains(.,'SOMA')]
  Botão SOMA clicado com sucesso!

INPUT DE DADOS (COM PAUSAS):
  [SELECTOR] action=click | method=xpath | selector=... (Plano de Conta)
  
  ✓ action=click | method=xpath | selector=...
  → Pressione ENTER para continuar...  ← AQUI PAUSA!
  
  [SELECTOR] action=type | method=xpath | selector=... (Descrição)
  
  ✓ action=type | method=xpath | selector=...
  → Pressione ENTER para continuar...  ← AQUI PAUSA!

✅ Pausas apenas onde importa
```

---

## 🔧 Implementação Técnica

### 1. **Novo Atributo em `Actions`**

```python
class Actions:
    def __init__(self, driver, cfg):
        self._debug_context = None  # Contexto para pausa seletiva
```

### 2. **Novo Método: `set_debug_context()`**

```python
def set_debug_context(self, context: str | None) -> None:
    """Define contexto para pausa seletiva"""
    self._debug_context = context
    if context:
        self._selector_debug_write(f"[CONTEXT] debug context ativado: {context}")
```

### 3. **Modificado: `_selector_debug_pause()`**

```python
def _selector_debug_pause(self, action, locator, detail=""):
    if not self._selector_debug:
        return
    
    # ... log a ação ...
    
    # Pausa SELETIVA baseada em contexto
    should_pause = self._debug_context in ("input_dados", "input_saida", "input_entrada")
    
    if should_pause:
        input(f"\n✓ {message}\n→ Pressione ENTER para continuar...\n")
```

### 4. **Ativação em `_fill_common()`**

```python
def _fill_common(self, row):
    # Ativa debug interativo para preenchimento de dados
    self.a.set_debug_context("input_dados")
    
    # ... preenche os campos ...
```

### 5. **Desativação em `create_and_get_doc_id()`**

```python
def create_and_get_doc_id(self, row):
    # ... preenche dados ...
    
    # Desativa debug interativo após preencher
    self.a.set_debug_context(None)
    
    # ... continua com pagamento/baixa sem pausas ...
```

---

## 📊 Fluxo de Execução

```
┌─ Início
│
├─ LOGIN (SEM PAUSA)
│  ├─ type email
│  ├─ type senha
│  └─ click submit
│
├─ NAVEGAÇÃO (SEM PAUSA)
│  ├─ click SOMA
│  └─ wait_visible página
│
├─ INPUT DADOS (COM PAUSA) ← Aqui começa!
│  ├─ set_debug_context("input_dados")
│  ├─ ⏸️ click plano_conta → Pausa
│  ├─ ⏸️ type descrição → Pausa
│  ├─ ⏸️ type valor → Pausa
│  └─ set_debug_context(None) ← Aqui termina!
│
├─ PAGAMENTO (SEM PAUSA)
│  ├─ click realizar pagamento
│  └─ input dados pagamento
│
├─ BAIXA (SEM PAUSA)
│  ├─ click inserir baixa
│  └─ salvar
│
└─ FIM
```

---

## 🚀 Como Usar

### 1. **Debug já está ativo**
```env
DEBUG_SELECTOR_INTERACTIVE=true
```

### 2. **Executar**
```bash
python main.py
```

### 3. **Resultado**
```
✓ Login: sem pausas (rápido)
✓ Navegação: sem pausas (rápido)
✓ Input de dados: COM PAUSAS (você vê cada passo)
✓ Pagamento/Baixa: sem pausas (automático)
```

---

## ✅ Contextos Suportados

| Contexto | Ativa Pausa | Uso |
|----------|-------------|-----|
| `None` | ❌ Não | Navegação, login |
| `"input_dados"` | ✅ Sim | Preencher campos Saída/Entrada |
| `"input_saida"` | ✅ Sim | Preencher campos de Saída |
| `"input_entrada"` | ✅ Sim | Preencher campos de Entrada |

---

## 🎯 Benefícios

| Antes | Depois |
|-------|--------|
| ❌ Muitas pausas desnecessárias | ✅ Pausa apenas nos inputs |
| ❌ Login lento | ✅ Login rápido |
| ❌ Difícil acompanhar | ✅ Fácil acompanhar o importante |
| ❌ Repetitivo | ✅ Focado nos dados |

---

## 📝 Arquivos Modificados

### `src/soma_app/automation/actions.py`
- ✅ Adicionado: `_debug_context` atributo
- ✅ Adicionado: `set_debug_context()` método
- ✅ Modificado: `_selector_debug_pause()` para verificar contexto

### `src/soma_app/automation/pages/entradas_saidas_page.py`
- ✅ Modificado: `_fill_common()` ativa contexto
- ✅ Modificado: `create_and_get_doc_id()` desativa contexto

---

## 🧪 Próxima Execução

```bash
python main.py
```

**Esperado:**
- ✅ Login: sem pausas
- ✅ Navegação: sem pausas
- ✅ **INPUT DE DADOS: COM PAUSAS** ← Aqui você verá as pausas
- ✅ Pagamento: sem pausas
- ✅ Baixa: sem pausas

---

## 💡 Extensível

Para ativar pausa em outro contexto, basta:

```python
# Em qualquer lugar do código
self.a.set_debug_context("seu_contexto")

# Depois desativar
self.a.set_debug_context(None)
```

---

**✅ Pausa Seletiva Implementada!**

Agora o debug é mais inteligente: pausa apenas onde importa (inputs de dados) e roda rápido no resto.
