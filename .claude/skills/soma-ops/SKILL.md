---
name: soma-ops
description: Operate and debug the SOMA automation project — running workflows, toggling env vars, reading logs/artifacts, and adjusting Selenium locators. Use when the user asks to run, debug, or troubleshoot the SOMA automation (e.g. "roda o SOMA", "por que o login falhou", "muda um seletor").
---

# SOMA Ops

Operational playbook for running and debugging the SOMA automation (`src/soma_app/`). See `AGENTS.md` at the repo root for architecture.

## Running

```bash
# one-shot run
.venv\Scripts\python main.py

# loop (scheduler)
.venv\Scripts\python agendador.py

# equivalent entrypoints (after pip install -e .)
soma-run
soma-scheduler
```

Config comes from `deploy/.env` by default. Point to a different file with `ENV_FILE=path/to/.env`.

## Key env var toggles

| Var | Default | Effect |
|---|---|---|
| `HEADLESS` | `true` | Chrome headless mode |
| `SOMA_BACKEND` | `selenium` | `api` uses REST client (`SomaApiClient`); `selenium` uses Page Objects |
| `API_FALLBACK_SELENIUM` | `true` | When `SOMA_BACKEND=api`, fall back to Selenium on API auth failure |
| `ALLOW_RETRY_ERROR` | `false` | Reprocess rows already marked `EM ERRO` |
| `RUN_CAIXAS_BANCOS` | `true` | Run the Caixas/Bancos post-process after the batch loop |
| `BATCH_SIZE` | `20` | Rows locked per preprocess batch |
| `SHEET_CONTAORDEM` / `SHEET_NAME` / `SHEET` | `TESTE_CONTAORDEM` | Worksheet name for the row queue |
| `SOMA_API_BASE_URL` / `SOMA_API_LOGIN` / `SOMA_API_PASSWORD` / `SOMA_SESSION_TOKEN` | — | API backend credentials |

All are parsed via `soma_app.infra.env` (`env_bool`/`env_str`/`env_int`) — check that module for exact parsing rules if a toggle isn't behaving as expected.

## Logs & artifacts

- Logs configured in `soma_app/infra/log_config.py`: root `APP` logger, `soma_audit`, and `soma_report` (legacy, no timestamps).
- `LOG_DIR` (default `logs/`) and `SCREENSHOTS_DIR` (default `artifacts/screenshots`) from `Settings`.
- Each run gets a `RUN_ID` (see `soma_app.infra.trace.new_run_id`); `[STEP]/[STEP_OK]/[STEP_FAIL]` lines in the log bracket major phases (`run.preprocess`, `run.init`, `run.login_ui`, `run.process_row`, `run.post.*`).
- ChromeDriver/Chrome version mismatches are logged at `run.init` via `get_chromedriver_info`/`get_chrome_version` (`soma_app/infra/webdriver_factory.py`) — check this line first when Selenium fails to start.

## Adjusting Selenium locators

If the SOMA website changes a selector, prefer overriding it in `config/locators.json` over editing Python:

1. Find the Page Object in `src/soma_app/automation/pages/*_page.py` and the class-attribute locator name.
2. Add/edit an entry in `config/locators.json` (loaded via `soma_app/config/locators.py::load_page_locator_config` + `apply_locator_overrides`).
3. Re-run with `HEADLESS=false` to visually confirm the new locator resolves.

Only edit the Page Object Python directly if the change is structural (new page, new flow) rather than a selector value.

## Debugging a failed row

1. Find the row's `run_id`/`row` in the log via the `[STEP]` lines for `run.process_row`.
2. Check the sheet's `STATUS`/`DOC. SOMA` columns — `EM ERRO` + `PENDENTE_DOC` status means `_should_recover_doc` will retry via `recover_doc_id` next run; plain `EM ERRO` needs `ALLOW_RETRY_ERROR=true` to be retried as a fresh create.
3. For API-backend failures, check for `EntradasSaidasAdapter`/`TransferenciasAdapter` fallback warnings (`soma_app/automation/adapters.py`) — a 401/403 disables the API client for the rest of the run and falls back to Selenium.

## Tests & lint before committing

```bash
.venv\Scripts\python -m pytest -q
.venv\Scripts\python -m ruff check .
```
