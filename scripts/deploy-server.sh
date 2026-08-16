#!/usr/bin/env bash

set -Eeuo pipefail

PROJECT_DIR="${1:-/home/ubuntu/soma-automation/SOMA}"
PROCESS_NAME="soma-automation"

log() {
  printf '%s\n' "$*"
}

fail() {
  log "ERRO: $*"
  exit 1
}

log "================================================"
log "DEPLOYMENT SEGURO NO SERVIDOR - SOMA"
log "================================================"

cd "$PROJECT_DIR" || fail "Diretório não encontrado: $PROJECT_DIR"
log "Diretório: $(pwd)"

command -v git >/dev/null 2>&1 || fail "git não está instalado."
command -v pm2 >/dev/null 2>&1 || fail "pm2 não está instalado."

if [ -n "$(git status --porcelain --untracked-files=no)" ]; then
  git status --short
  fail "Existem alterações locais rastreadas no servidor. Deployment interrompido para não perder dados."
fi

log "[1/6] Atualizando referências do GitHub..."
git fetch origin main --prune
REMOTE_SHA="$(git rev-parse origin/main)"
LOCAL_SHA="$(git rev-parse HEAD)"
log "Commit servidor antes: $LOCAL_SHA"
log "Commit GitHub:         $REMOTE_SHA"

log "[2/6] Atualizando a branch main por fast-forward..."
git checkout main
git pull --ff-only origin main

[ "$(git rev-parse HEAD)" = "$(git rev-parse origin/main)" ] \
  || fail "O servidor não ficou no mesmo commit do GitHub."

log "[3/6] Preparando ambiente Python..."
if [ ! -x ".venv/bin/python" ]; then
  python3 -m venv .venv
fi

.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.lock.txt

log "[4/6] Validando o código..."
.venv/bin/python -m compileall -q src main.py agendador.py
mkdir -p logs artifacts/screenshots artifacts/diagnostics

log "[5/6] Iniciando ou recarregando o job automático..."
pm2 startOrReload ecosystem.config.js --only "$PROCESS_NAME" --update-env
pm2 save

log "[6/6] Validando status e logs..."
pm2 status "$PROCESS_NAME"
pm2 logs "$PROCESS_NAME" --lines 50 --nostream

FINAL_SHA="$(git rev-parse HEAD)"
log "================================================"
log "DEPLOYMENT CONCLUÍDO"
log "Commit ativo: $FINAL_SHA"
log "Processo PM2: $PROCESS_NAME"
log "================================================"
