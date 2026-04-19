# SOMA Automation

Projeto de automacao financeira para o sistema SOMA, com leitura/escrita em Google Sheets e execucao por Selenium/API.

## O que foi padronizado

- Estrutura de pacote Python em `src/soma_app`
- Entradas unificadas:
  - `python main.py` -> executa o workflow principal
  - `python agendador.py` -> executa em loop
- Configuracao centralizada em variaveis de ambiente (`.env`)
- Base de qualidade com `pyproject.toml`, `ruff` e `pytest`

## Estrutura

```text
SOMA/
  config/                  # compatibilidade com scripts legados
  src/soma_app/
    automation/            # pages, actions, API adapter
    config/                # settings
    domain/                # modelos e regras
    infra/                 # webdriver, sheets client, logging, tracing
    workflows/             # orquestracao principal
    scheduler.py           # loop de execucao
  tests/                   # testes unitarios
  main.py                  # wrapper para workflow principal
  agendador.py             # wrapper para scheduler
  pyproject.toml
  .env.example
```

## Requisitos

- Python 3.12+
- Google Chrome instalado
- Conta de servico Google com acesso a planilha

## Setup rapido

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
# opcional (ferramentas de desenvolvimento)
pip install -r requirements-dev.txt
```

Para instalacao reprodutivel com versoes fixas:

```bash
pip install -r requirements.lock.txt
```

Crie o arquivo de ambiente:

```bash
copy .env.example deploy\.env
```

Preencha no `deploy/.env`:

- `GOOGLE_CREDENTIALS_PATH`
- `SPREADSHEET_URL`
- `SITE_USER`
- `SITE_PASSWORD`

## Execucao

Rodar uma vez:

```bash
python main.py
```

Rodar em loop:

```bash
python agendador.py
```

Tambem disponivel por entrypoint:

```bash
soma-run
soma-scheduler
```

## Qualidade

```bash
ruff check .
pytest
```

## Seguranca

- Nao versionar credenciais, tokens e arquivos `.env`
- Manter `deploy/credenciais.json` e `deploy/.env` apenas localmente
