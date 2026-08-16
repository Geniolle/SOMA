#!/usr/bin/env bash

set -Eeuo pipefail

PROCESS_NAME="soma-automation"
PROJECT_DIR="${SOMA_PROJECT_DIR:-/home/ubuntu/soma-automation/SOMA}"

log() {
  printf '%s\n' "$*"
}

fail() {
  log "ERRO: $*"
  exit 1
}

ensure_pm2() {
  command -v pm2 >/dev/null 2>&1 || fail "pm2 não está instalado."
}

ensure_project_dir() {
  cd "$PROJECT_DIR" || fail "Diretório não encontrado: $PROJECT_DIR"
}

pm2_start_or_reload() {
  ensure_project_dir
  pm2 startOrReload ecosystem.config.js --only "$PROCESS_NAME" --update-env
  pm2 save
}

check_chrome() {
  if command -v google-chrome >/dev/null 2>&1; then
    google-chrome --version
    return 0
  fi

  if command -v chromium-browser >/dev/null 2>&1; then
    chromium-browser --version
    return 0
  fi

  return 1
}

check_chromedriver() {
  if command -v chromedriver >/dev/null 2>&1; then
    chromedriver --version
    return 0
  fi

  if command -v chromium-chromedriver >/dev/null 2>&1; then
    chromium-chromedriver --version
    return 0
  fi

  return 1
}

usage() {
  cat <<'EOF'
Uso: ./control-soma.sh <comando>

Comandos:
  check    Verificar Chrome, ChromeDriver, PM2 e iniciar/reiniciar o job
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

ensure_pm2

COMMAND="${1:-}"
[ -n "$COMMAND" ] || usage

case "$COMMAND" in
  check)
    log "================================================"
    log "Teste SOMA - Verificar Chrome e Iniciar"
    log "================================================"
    log ""

    log "1️⃣ Verificando Chrome..."
    if CHROME_VERSION="$(check_chrome)"; then
      log "   ✅ $CHROME_VERSION"
    else
      fail "Chrome não encontrado!"
    fi
    log ""

    log "2️⃣ Verificando ChromeDriver..."
    if CHROMEDRIVER_VERSION="$(check_chromedriver)"; then
      log "   ✅ $CHROMEDRIVER_VERSION"
    else
      log "   ❌ ChromeDriver não encontrado!"
      log "   Instalando..."
      sudo apt-get install -y chromium-chromedriver
      CHROMEDRIVER_VERSION="$(chromedriver --version)"
      log "   ✅ $CHROMEDRIVER_VERSION"
    fi
    log ""

    log "3️⃣ Verificando PM2..."
    log "   ✅ PM2 encontrado"
    log ""

    log "4️⃣ Iniciando/Reiniciando SOMA..."
    pm2_start_or_reload
    sleep 2
    log ""

    log "5️⃣ Status Final:"
    pm2 status
    log ""

    log "6️⃣ Logs SOMA (últimas 20 linhas):"
    log "---"
    pm2 logs "$PROCESS_NAME" --lines 20 --nostream 2>/dev/null || log "Nenhum log ainda"
    log "---"
    log ""

    log "================================================"
    log "✅ TESTE COMPLETADO!"
    log "================================================"
    log ""
    log "Comandos úteis:"
    log "   pm2 logs soma-automation         # Ver logs"
    log "   pm2 logs soma-automation logs-f  # Tempo real"
    log "   pm2 monit                        # Dashboard"
    log "   bash control-soma.sh health      # Health check"
    log ""
    ;;
  status)
    pm2 status "$PROCESS_NAME"
    ;;
  start)
    pm2_start_or_reload
    pm2 status "$PROCESS_NAME"
    ;;
  stop)
    pm2 stop "$PROCESS_NAME"
    pm2 save
    ;;
  restart)
    pm2_start_or_reload
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
    ensure_project_dir
    echo "Commit ativo: $(git rev-parse --short HEAD 2>/dev/null || echo desconhecido)"
    echo "Commit GitHub: $(git rev-parse --short origin/main 2>/dev/null || echo desconhecido)"
    pm2 status "$PROCESS_NAME"
    pm2 logs "$PROCESS_NAME" --lines 30 --nostream
    ;;
  *)
    usage
    ;;
esac
