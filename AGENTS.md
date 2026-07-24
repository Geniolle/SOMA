# AGENTS.md — SOMA Automation

Guia para agentes de IA (e humanos) trabalhando neste repositório. Automação financeira do sistema SOMA: lê/escreve linhas numa Google Sheet e executa lançamentos via Selenium (UI) e/ou API REST.

## Arquitetura (camadas)

```
src/soma_app/
  domain/       # regras de negócio puras, sem IO (models.py, rules.py)
  infra/        # IO: webdriver, sheets_client, soma_api_client, logging, tracing, env
  automation/   # Selenium Page Objects (automation/pages/*), API adapters (automation/api/*)
                # e adapters.py (fallback API->Selenium: EntradasSaidasAdapter/TransferenciasAdapter)
  workflows/    # orquestração: run_soma.py (entrypoint principal), process_*.py, scheduler.py
  config/       # Settings (dataclass, .from_env()) e locators.py (override de seletores)
```

Regra de dependência: `domain` não importa de `infra`/`automation`/`workflows`. Helpers de IO (env vars, webdriver) vivem em `infra/`; regras de negócio puras (ex.: `is_error_doc`, `should_process`) vivem em `domain/rules.py`.

## Convenções de variáveis de ambiente

- Ler env vars sempre via `soma_app.infra.env` (`env_bool`, `env_str`, `env_int`) — não duplicar parsing de `os.getenv` em módulos novos.
- `Settings` (`src/soma_app/config/settings.py`) é a fonte de verdade para configuração "estrutural" (credenciais, planilha, timeouts). Toggles de execução pontual (`HEADLESS`, `SOMA_BACKEND`, `ALLOW_RETRY_ERROR`, etc.) são lidos diretamente via `infra.env` dentro de `run_soma.py`.
- `.env` é carregado de `deploy/.env` por padrão, ou do caminho em `ENV_FILE`.

## Fluxo principal (`workflows/run_soma.py`)

1. `_load_settings()` + `configure_logging()` + `preprocess_contaordem()` (lock de linhas na sheet).
2. Se `SOMA_BACKEND=api`: usa `EntradasSaidasApi`/`TransferenciasApi`, com fallback opcional para Selenium via `_EntradasSaidasAdapter`/`_TransferenciasAdapter` (`API_FALLBACK_SELENIUM=true`).
3. Se `SOMA_BACKEND=selenium` (default): usa `automation/pages/*_page.py` diretamente.
4. Loop de batches: processa `result.workset`, grava OK/erro via `_mark_row_ok`/`_mark_row_error`, e roda pós-processos (`process_caixas_bancos`, `process_soma`) quando a fila esvazia.

## Testes e qualidade

```bash
.venv\Scripts\python -m pytest -q
.venv\Scripts\python -m ruff check .
```

Rodar sempre antes de propor uma mudança em `workflows/` ou `infra/webdriver_factory.py` — são os pontos de maior acoplamento.

## Locators (seletores Selenium)

Seletores de Page Objects (`automation/pages/*_page.py`) podem ser sobrescritos em runtime por `config/locators.json`, processado por `soma_app/config/locators.py` (`load_page_locator_config` + `apply_locator_overrides`). Preferir editar o JSON a alterar o Python quando o site mudar um seletor.

## Segurança — nunca commitar

- `deploy/.env`, `deploy/credenciais.json`, `chave.key` (já no `.gitignore`; confirmar antes de qualquer `git add`).
- Deploy em produção é via SSH manual (`ssh -i chave.key ubuntu@<host>` + `git reset --hard origin/main`) — não há CI/CD automatizado. Qualquer sugestão de deploy deve ser confirmada explicitamente com o usuário antes de executar.

## Ao propor refatorações

- Não assumir que funções com o mesmo nome em módulos diferentes (`_now_pt`, `_norm`, etc.) fazem a mesma coisa — ler a implementação antes de consolidar.
- `run_soma.py::main()` é só orquestração: delega para `_bootstrap_backend` (browser/API), `_build_processors`, `_run_batches` (loop de batches + `_process_row` por linha) e `_build_summary`. Estado que precisa sobreviver a exceções parciais (ex.: `bundle`/`api_client` criados a meio do bootstrap) usa `RunState`/`RunTotals`, não variáveis soltas — preserva o cleanup no `finally` de `main()` mesmo se `_bootstrap_backend` falhar no meio. Qualquer nova extração deve manter a ordem exata dos `with step(logger, ...)` (rastreabilidade em produção via `run_id`/`RUN_ID`).
