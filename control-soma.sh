#!/usr/bin/env bash

set -Eeuo pipefail

PROCESS_NAME="soma-automation"
PROJECT_DIR="${SOMA_PROJECT_DIR:-/home/ubuntu/soma-automation/SOMA}"

usage() {
  cat <<'EOF'
Uso: ./control-soma.sh <comando>

Comandos:
  status   Ver o estado do job
  start    Iniciar o job automático pelo ecosystem.config.js
  stop     Parar o job
  restart  Recarregar o job e as variáveis de ambiente
  logs     Mostrar os últimos 50 registos
  logs-f   Acompanhar logs em tempo real
  monit    Abrir o monitor do PM2
  health   Verificar processo, commit e logs recentes
EOF
  exit 1
}

command -v pm2 >/dev/null 2>&1 || {
  echo "ERRO: pm2 não está instalado."
  exit 1
}

COMMAND="${1:-}"
[ -n "$COMMAND" ] || usage

case "$COMMAND" in
  status)
    pm2 status "$PROCESS_NAME"
    ;;
  start)
    cd "$PROJECT_DIR"
    pm2 startOrReload ecosystem.config.js --only "$PROCESS_NAME" --update-env
    pm2 save
    pm2 status "$PROCESS_NAME"
    ;;
  stop)
    pm2 stop "$PROCESS_NAME"
    pm2 save
    ;;
  restart)
    cd "$PROJECT_DIR"
    pm2 startOrReload ecosystem.config.js --only "$PROCESS_NAME" --update-env
    pm2 save
    pm2 status "$PROCESS_NAME"
    ;;
  logs)
    pm2 logs "$PROCESS_NAME" --lines 50 --nostream
    ;;
  logs-f)
    pm2 logs "$PROCESS_NAME"
    ;;
  monit)
    pm2 monit
    ;;
  health)
    cd "$PROJECT_DIR"
    echo "Commit ativo: $(git rev-parse --short HEAD 2>/dev/null || echo desconhecido)"
    echo "Commit GitHub: $(git rev-parse --short origin/main 2>/dev/null || echo desconhecido)"
    pm2 status "$PROCESS_NAME"
    pm2 logs "$PROCESS_NAME" --lines 30 --nostream
    ;;
  *)
    usage
    ;;
esac
