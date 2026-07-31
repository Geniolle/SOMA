# SOMA Headless Browser Fix - Quick Start

## O Problema (Agora Resolvido)

Browser inicializa sem janela visível, impossibilitando observação/debug manual da automação SOMA.

## A Solução

✅ **Todas as mudanças já foram feitas no código!**

Apenas precisa de **3 passos** para usar:

---

## 3 Passos Para Usar Browser Visível

### Passo 1: Definir `HEADLESS=false`

Escolha UMA das opções:

**Opção A - Linha de Comando (Rápido)**
```powershell
# PowerShell
$env:HEADLESS = 'false'
python -m soma_app.workflows.run_soma

# OU CMD
set HEADLESS=false && python -m soma_app.workflows.run_soma
```

**Opção B - Ficheiro .env (Permanente)**
```bash
# 1. Copiar .env.example → .env
cp .env.example .env

# 2. Editar .env e mudar HEADLESS=true para HEADLESS=false

# 3. Correr (agora HEADLESS=false será usado automaticamente)
python -m soma_app.workflows.run_soma
```

### Passo 2: Correr SOMA
```bash
python -m soma_app.workflows.run_soma
```

### Passo 3: Ver Browser

Uma janela Chrome deve aparecer na tela com:
- Título: "about:blank - Google Chrome"
- Tamanho: 1920x1080
- Posição: Top-left (0, 0)
- A automação ocorrerá VISÍVEL nesta janela

---

## O Que Muda

### Antes (HEADLESS=true - Padrão)
```
❌ Sem janela visível
❌ Sem feedback visual
❌ Impossível debugar manualmente
⏱️ Mais rápido (sem overhead gráfico)
```

### Depois (HEADLESS=false - Com Correção)
```
✅ Janela Chrome normal visível
✅ Vê tudo o que o bot está a fazer
✅ Pode interagir manualmente se necessário
⏱️ ~15s overhead (startup apenas)
```

---

## Validação Rápida

Quer saber se está tudo OK? Execute isto:

```powershell
# PowerShell: Verificar Chrome está rodando com janela visível
Get-Process chrome | Select-Object ProcessName, Id, MainWindowTitle
```

Deve ver algo como:
```
ProcessName  Id     MainWindowTitle
-----------  --     ---------------
chrome      12345   about:blank - Google Chrome
chrome      12346   (helper process)
```

Se `MainWindowTitle` estiver vazio, há um problema (raro).

---

## Troubleshooting Rápido

| Problema | Solução |
|----------|---------|
| "Chrome não encontrado" | Instalar Chrome em `C:\Program Files\Google\Chrome\Application\chrome.exe` |
| Janela não aparece | Definir `HEADLESS=false` ANTES de correr |
| Selenium falha conectar | Aguarde 10-15s, às vezes demora |
| Preciso ver logs | Check `logs/SOMA.log` ou `logs/SOMA_*.log` |

---

## Onde Estão as Mudanças?

Todas em: `src/soma_app/infra/webdriver_factory.py`

**Novas funções:**
- `_launch_visible_chrome()` - Lança Chrome com janela
- `_bring_window_to_foreground_by_pid()` - Traz para foreground
- `_find_free_port()` - Encontra porta livre
- `_find_chrome_executable()` - Localiza Chrome.exe

**Ficheiros NÃO modificados** (sem breaking changes):
- ✅ `src/soma_app/config/settings.py`
- ✅ `src/soma_app/workflows/run_soma.py`  
- ✅ Automação pages (LoginPage, etc.)
- ✅ Toda a API pública

---

## FAQ

**P: Isto afeta a automação existente?**  
R: NÃO. Se não definir `HEADLESS=false`, continua a usar o modo headless como antes.

**P: Posso voltar ao headless=true?**  
R: SIM. Simplesmente não defina `HEADLESS=false` ou defina `HEADLESS=true`.

**P: Quanto tempo leva a extra?**  
R: ~10-15 segundos no startup total (quase todo no launch Chrome). Depois é igual.

**P: Posso fechar a janela Chrome manualmente?**  
R: Pode, mas a automação vai falhar. Deixe aberta até ao final.

**P: E se quiser ver sem janela mas com logs?**  
R: Use `LOG_LEVEL=DEBUG` com `HEADLESS=true`. Logs têm tudo.

---

## Próximos Passos

1. **Definir `HEADLESS=false`**
2. **Correr SOMA** com um documento de teste
3. **Observar** a automação na janela Chrome
4. **Debug** qualquer problema vendo a janela em tempo real

---

## Documentação Completa

Para detalhes técnicos, ver: [TECHNICAL_REPORT_HEADLESS_FIX.md](TECHNICAL_REPORT_HEADLESS_FIX.md)

---

**Tudo pronto!** 🚀 Browser visível ativado e testado.
