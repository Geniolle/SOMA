# Modo Debug Interativo de Seletores

## Ativação

O modo debug interativo está configurado na variável de ambiente:

```env
DEBUG_SELECTOR_INTERACTIVE=true
```

**Localização:** `deploy/.env`

## Funcionamento

Quando `DEBUG_SELECTOR_INTERACTIVE=true`:

1. **Cada ação é logada** - cliques, inputs, seleções, etc. mostram:
   - `action`: tipo de ação (click, type, select_by_text, etc)
   - `method`: tipo de seletor (xpath, css selector, id, name, etc)
   - `selector`: o seletor/xpath completo
   - Detalhes adicionais (valor inserido, etc)

2. **Pausa após cada ação** - o script para e aguarda pressionar ENTER no terminal

3. **Logs salvos** - todas as ações são registradas em arquivo:
   - Arquivo: `logs/soma_selectors_YYYYMMDD_HHMMSS.log`
   - Console: mensagens prefixadas com `[SELECTOR]`

## Exemplo de Saída

```
[SELECTOR] START | modo interativo de seletores ativo

✓ action=type | method=name | selector=email | clear=True | value_length=27
→ Pressione ENTER para continuar...
```

*Usuario pressiona ENTER*

```
[SELECTOR] action=type | method=name | selector=email | clear=True | value_length=27

✓ action=type | method=name | selector=senha | clear=True | value_length=10
→ Pressione ENTER para continuar...
```

## Quando Usar

- 🔍 **Depuração de seletores**: verificar se o xpath/CSS está correto
- 🐛 **Troubleshooting de falhas**: pausar na ação problemática
- ✅ **Validação manual**: confirmar visualmente cada passo
- 📝 **Documentação visual**: revisar o fluxo completo do script

## Desativação

Para desativar, altere em `deploy/.env`:

```env
DEBUG_SELECTOR_INTERACTIVE=false
```

O script voltará a executar sem pausas e continuará logando as ações normalmente.

## Arquivos de Log

Os logs do modo interativo ficam em:
- **Seletores**: `logs/soma_selectors_*.log`
- **Aplicação**: `logs/soma_dev_*.log`
- **Auditoria**: `logs/soma_audit_*.log`
