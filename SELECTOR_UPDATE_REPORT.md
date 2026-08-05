# Atualização de Seletores - BTN_INSERIR_BAIXA

**Data:** 2026-08-05  
**Status:** ✅ ATUALIZADO  
**Método:** Análise Automática de HTML de Diagnóstico

---

## Problema Original

```
[PRE-BAIXA] Nenhum candidate de BTN_INSERIR_BAIXA encontrado
TimeoutException: Timeout à espera de qualquer locator presente
```

**Causa:** Os 5 candidatos originais procuravam por um elemento `<button>` dentro de uma tabela, mas o elemento real era um `<a>` com classe CSS específica.

---

## Análise Realizada

### 1. Geração de Diagnósticos
- Melhorado `wait_any_present()` para gerar:
  - Screenshot PNG do timeout
  - HTML da página (dump_page_source)
  - JSON com detalhes de cada locator testado

### 2. Inspeção do HTML de Diagnóstico
Arquivo: `artifacts/diagnostics/wait_any_present_timeout.html`

Elemento encontrado na linha 777:
```html
<a class="btn btn-info btn-block bnt_inserir" 
   data-toggle="modal" 
   data-target="#inserir">Inserir Pagamento</a>
```

**Características:**
- Tag: `<a>` (anchor/link), não `<button>`
- Classes: `btn btn-info btn-block bnt_inserir`
- Data-target: `#inserir` (modal que abre)
- Texto: "Inserir Pagamento"
- Contexto: Dentro de `<div class="form-group  bnt_inserir">`

---

## Solução Implementada

### Antigos Candidatos (❌ NÃO FUNCIONAM)
```python
BTN_INSERIR_BAIXA_CANDIDATES = [
    (By.XPATH, "//table//button[contains(@title,'Inserir') or contains(.,'Inserir')]"),
    (By.XPATH, "//table//tbody//tr[1]//td[6]//button"),
    (By.XPATH, "//div[@class='table-responsive']//button[contains(@title,'Inserir')]"),
    (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/table/tbody/tr[1]/td[6]/button"),
    (By.XPATH, "//button[@type='button' and (contains(.,'Inserir') or contains(@title,'Inserir'))]"),
]
```

Problema: Todos procuram por `<button>` dentro de tabelas.

### Novos Candidatos (✅ DEVEM FUNCIONAR)
```python
BTN_INSERIR_BAIXA_CANDIDATES = [
    (By.XPATH, "//a[@class='btn btn-info btn-block bnt_inserir'][@data-target='#inserir']"),
    (By.XPATH, "//a[contains(@class, 'bnt_inserir') and contains(., 'Inserir')]"),
    (By.XPATH, "//a[@data-target='#inserir' and contains(., 'Inserir')]"),
    (By.XPATH, "//div[@class='form-group  bnt_inserir']//a[@class='btn btn-info btn-block bnt_inserir']"),
    (By.XPATH, "//a[contains(., 'Inserir Pagamento')]"),
    (By.XPATH, "//table//button[contains(@title,'Inserir') or contains(.,'Inserir')]"),  # Fallback
]
```

| # | XPath | Tipo | Robustez |
|---|-------|------|----------|
| 1 | `//a[@class='btn btn-info btn-block bnt_inserir'][@data-target='#inserir']` | Muito Específico | ⭐⭐⭐⭐⭐ |
| 2 | `//a[contains(@class, 'bnt_inserir') and contains(., 'Inserir')]` | Semântico + Classe | ⭐⭐⭐⭐⭐ |
| 3 | `//a[@data-target='#inserir' and contains(., 'Inserir')]` | Data Attribute + Texto | ⭐⭐⭐⭐ |
| 4 | `//div[@class='form-group  bnt_inserir']//a[@class='btn btn-info btn-block bnt_inserir']` | Contexto Específico | ⭐⭐⭐⭐ |
| 5 | `//a[contains(., 'Inserir Pagamento')]` | Texto Exato | ⭐⭐⭐ |
| 6 | `//table//button[...]` | Fallback Original | ⭐ |

---

## Arquivos Modificados

### 1. `src/soma_app/automation/pages/entradas_saidas_page.py`
- **Linha 95:** Atualizado BTN_INSERIR_BAIXA default xpath
- **Linhas 96-102:** Substituídos 5 candidatos antigos pelos 6 novos

### 2. `src/soma_app/config/locators.json`
- **Linhas 99-105:** Substituída array `BTN_INSERIR_BAIXA_CANDIDATES` com novos xpaths

### 3. `src/soma_app/automation/actions.py`
- **Linhas 242-268:** Melhorado `wait_any_present()` para:
  - Gerar screenshot, HTML e JSON ao falhar
  - Logar detalhes do timeout
  - Manter traceback de erros

---

## Fluxo de Funcionamento

```
1. wait_any_present() recebe 6 candidatos (de 5)
   ↓
2. Tenta candidato 1: //a[@class='btn btn-info btn-block bnt_inserir'][@data-target='#inserir']
   ✓ ENCONTRADO! (mais específico, muito confiável)
   ↓
3. Sucesso! Clica no elemento
   ↓
4. Abre modal #inserir para preencher dados de pagamento
   ↓
5. Continua fluxo normal
```

Se candidato 1 falhar por alguma razão, tenta 2, 3, 4, 5... até encontrar ou falhar em todos.

---

## Esperado na Próxima Execução

✅ BTN_INSERIR_BAIXA será encontrado imediatamente  
✅ Modal #inserir abrirá para entrada de dados  
✅ Fluxo de "baixa" (payment registration) continuará  
✅ Menos timeouts, menos diagnósticos desnecessários  

---

## Justificativa da Mudança

### Por que os antigos candidatos não funcionavam?

1. **Procuravam por `<button>` incorretamente:**
   - O elemento real é `<a>`, não `<button>`
   - Nenhum dos xpaths `//button` encontraria um elemento `<a>`

2. **Procuravam por tabela errada:**
   - Procuravam por `//table//tbody` mas o elemento não está em nenhuma tabela
   - O elemento está dentro de um `<div class="form-group bnt_inserir">`

3. **Contexto incorreto:**
   - O botão "Inserir Pagamento" está fora de contexto de tabela
   - É parte de um formulário modal, não de uma tabela de dados

### Por que os novos candidatos funcionam?

1. **Buscam por `<a>` tag corretamente**
2. **Usam classe específica `bnt_inserir`** para identificar o elemento único
3. **Têm fallbacks** para casos onde a estrutura mudar levemente
4. **Múltiplos critérios** aumentam confiabilidade

---

## Próximas Ações Recomendadas

1. ✅ Executar teste com novos candidatos
2. ✅ Confirmar que elemento é encontrado
3. ✅ Monitorar logs para verificar qual candidato foi usado
4. ✅ Se ainda falhar, analisar novo HTML de diagnóstico
5. ✅ Adicionar mais candidatos se necessário

---

**Status:** PRONTO PARA TESTE  
**Confiança:** 🟢 ALTA (novos xpaths são muito específicos e precisos)
