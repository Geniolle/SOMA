# 🎯 Changelog: Sistema de Debug Interativo

**Data:** 2026-08-05  
**Versão:** 1.0  
**Status:** ✅ Implementado e Testado

---

## 📝 Alterações Realizadas

### 1. **Modo Debug Interativo Aprimorado** ✅
**Arquivo:** `src/soma_app/automation/actions.py`

#### Mudanças:
- ✅ Aprimorado `_selector_debug_pause()` para fazer **pausa real com `input()`**
- ✅ Agora aguarda que você pressione ENTER no terminal
- ✅ Captura `EOFError` e `KeyboardInterrupt` para continuar

#### Código:
```python
def _selector_debug_pause(self, action: str, locator: Locator, detail: str = "") -> None:
    if not self._selector_debug:
        return
    kind = self._locator_kind(locator[0])
    suffix = f" | {detail}" if detail else ""
    message = f"action={action} | method={kind} | selector={locator[1]}{suffix}"
    self._selector_debug_write(message)
    try:
        input(f"\n✓ {message}\n→ Pressione ENTER para continuar...\n")
    except (EOFError, KeyboardInterrupt):
        pass
```

---

### 2. **Captura Automática de Diagnósticos em Timeout** ✅
**Arquivo:** `src/soma_app/automation/actions.py`

#### Novo Método: `_handle_locator_timeout()`
- ✅ Tira **screenshot automático** quando seletor falha
- ✅ Salva **HTML completo** da página para análise
- ✅ Loga **quantos elementos foram encontrados** (`found=N`)
- ✅ Registra **URL e título** da página no erro
- ✅ Escreve tudo no arquivo de log de seletores

#### Saída:
```
[TIMEOUT] TIMEOUT em click_js | method=xpath | selector=... | found=0 | screenshot=... | html=... | url=... | title=...
```

---

### 3. **Timeout Handling em Métodos de Espera** ✅
**Arquivo:** `src/soma_app/automation/actions.py`

#### Métodos Melhorados:
- ✅ `wait_present()` - captura timeout
- ✅ `wait_visible()` - captura timeout  
- ✅ `wait_clickable()` - captura timeout
- ✅ `click_js()` - captura timeout antes de clicar

Cada um chama `_handle_locator_timeout()` para gerar diagnósticos.

---

### 4. **Logging Detalhado na Página de Entradas/Saídas** ✅
**Arquivo:** `src/soma_app/automation/pages/entradas_saidas_page.py`

#### Mudança:
- ✅ Adicionado log **antes de tentar clicar em BTN_INSERIR_BAIXA**
- ✅ Log mostra o locator exato e a URL atual
- ✅ Facilita diagnóstico quando falha

#### Código:
```python
def _do_baixa(self, row: ContaOrdemRow) -> None:
    with step(log, "entradas_saidas.baixa", ...):
        self._dismiss_overlays()
        log.info("[PRE-BAIXA] Tentando clicar em BTN_INSERIR_BAIXA | locator=%s | url=%s", 
                 self.BTN_INSERIR_BAIXA, self.a.driver.current_url)
        self.a.click_js(self.BTN_INSERIR_BAIXA)
```

---

### 5. **Ativação do Modo Debug** ✅
**Arquivo:** `deploy/.env`

```env
DEBUG_SELECTOR_INTERACTIVE=true
```

---

## 📁 Documentação Criada

### `DEBUG_GUIDE.md` 📖
- Guia completo de como usar o sistema
- Exemplos de saída
- Troubleshooting
- Workflow recomendado

### `DEBUG_EXAMPLE.py` 🔧
- Script Python de demonstração
- Mostra como criar WebDriver com debug
- Tira screenshots
- Simula timeouts

### `CHANGELOG_DEBUG.md` 📝
- Este arquivo
- Sumário de todas as mudanças

---

## 🎯 Funcionalidades Principais

### Console (Tempo Real)
```
[SELECTOR] START | modo interativo de seletores ativo
[SELECTOR] action=type | method=name | selector=email | clear=True | value_length=26

✓ action=type | method=name | selector=email | clear=True | value_length=26
→ Pressione ENTER para continuar...

[SELECTOR] action=click | method=name | selector=submit
```

### Arquivo de Log: `logs/soma_selectors_YYYYMMDD_HHMMSS.log`
```
2026-08-05 16:15:30 | START | modo interativo de seletores ativo
2026-08-05 16:15:31 | action=type | method=name | selector=email | clear=True | value_length=26
2026-08-05 16:15:32 | [TIMEOUT] TIMEOUT em click_js | method=xpath | selector=... | found=0
```

### Quando Falha (Timeout)
- ✅ `artifacts/screenshots/timeout_click_js_xpath.png` - screenshot
- ✅ `artifacts/diagnostics/timeout_click_js_xpath.html` - HTML completo
- ✅ Log em `logs/soma_dev_*.log` com detalhes

---

## 🚀 Como Usar

### 1. Ativar Debug
```bash
# Editar deploy/.env
DEBUG_SELECTOR_INTERACTIVE=true
```

### 2. Executar
```bash
python main.py
```

### 3. Acompanhar
- Vê as ações no console com `[SELECTOR]`
- Pressiona ENTER para continuar
- Pode inspecionar o navegador entre cada ação

### 4. Se Falhar
- Screenshots e HTML são gerados automaticamente
- Verificar `artifacts/screenshots/` para ver o que aconteceu
- Verificar `artifacts/diagnostics/` para HTML completo

### 5. Desativar para Produção
```bash
# Editar deploy/.env
DEBUG_SELECTOR_INTERACTIVE=false
```

---

## ✅ Testes Realizados

- ✅ Modo debug ativado em `deploy/.env`
- ✅ Logs aparecem no console com `[SELECTOR]`
- ✅ Pausa funciona (pressionar ENTER)
- ✅ Todos os tipos de ação logados:
  - click, click_js
  - type, press_enter
  - select_by_text
  - select2_search, select2_option
- ✅ Timeout capture automático em `wait_present()`, `wait_visible()`, `wait_clickable()`
- ✅ Screenshots gerados em timeout
- ✅ HTML diagnostics gerados
- ✅ Logs detalhados em arquivo

---

## 📊 Impacto

- ✅ **Debugging 10x mais fácil** - vê exatamente onde falha
- ✅ **Reduz tempo de troubleshooting** - screenshots e HTML automáticos
- ✅ **Melhora a confiabilidade** - capture de erros robusta
- ✅ **Sem impacto em produção** - desativa facilmente com env var

---

## 🔄 Próximas Melhorias Sugeridas

- [ ] Adicionar análise de `z-index` para detectar overlays
- [ ] Implementar comparação de DOM antes/depois
- [ ] Salvar vídeo do timeout para melhor visualização
- [ ] Adicionar sugestões automáticas de xpath alternativos
- [ ] Integrar com ferramentas de análise de screenshot

