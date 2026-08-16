#!/usr/bin/env bash
set -u

PROJECT_DIR="/home/ubuntu/soma-automation/SOMA"
PYTHON_BIN="$PROJECT_DIR/.venv/bin/python"
MAIN_FILE="$PROJECT_DIR/main.py"
LOCK_FILE="/tmp/soma.lock"

NTFY_TOPIC="${NTFY_TOPIC:-soma-alerta-clayton}"
NTFY_URL="https://ntfy.sh/${NTFY_TOPIC}"

LOG_FILE="$PROJECT_DIR/logs/soma-run.log"
LAST_OK_FILE="$PROJECT_DIR/runtime/soma.last_ok"
LAST_ERROR_FILE="$PROJECT_DIR/runtime/soma.last_error"

mkdir -p "$PROJECT_DIR/logs" "$PROJECT_DIR/runtime"

send_alert() {
  local title="$1"
  local message="$2"
  local priority="${3:-5}"
  local tags="${4:-rotating_light}"

  curl -s \
    -H "Title: ${title}" \
    -H "Priority: ${priority}" \
    -H "Tags: ${tags}" \
    -d "${message}" \
    "$NTFY_URL" >/dev/null || true
}

log() {
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" >> "$LOG_FILE"
}

cd "$PROJECT_DIR" || {
  send_alert "SOMA error" "Could not access $PROJECT_DIR on $(hostname)."
  exit 1
}

if [ ! -x "$PYTHON_BIN" ]; then
  send_alert "SOMA error" "Virtualenv Python not found or not executable: $PYTHON_BIN"
  exit 1
fi

if [ ! -f "$MAIN_FILE" ]; then
  send_alert "SOMA error" "main.py not found at $MAIN_FILE"
  exit 1
fi

log "Starting SOMA execution..."

flock -n "$LOCK_FILE" "$PYTHON_BIN" "$MAIN_FILE" >> "$LOG_FILE" 2>&1
exit_code=$?

if [ "$exit_code" -eq 0 ]; then
  date '+%Y-%m-%d %H:%M:%S' > "$LAST_OK_FILE"
  log "SOMA execution completed successfully."
  exit 0
fi

if [ "$exit_code" -eq 1 ]; then
  log "Execution skipped: active lock, previous SOMA process still running."
  exit 0
fi

date '+%Y-%m-%d %H:%M:%S' > "$LAST_ERROR_FILE"
log "ERROR: SOMA execution failed with code $exit_code."

send_alert \
  "SOMA failed" \
  "SOMA finished with error on $(hostname). Code: $exit_code. Folder: $PROJECT_DIR. Check logs/soma-run.log." \
  "5" \
  "rotating_light"

exit "$exit_code"
