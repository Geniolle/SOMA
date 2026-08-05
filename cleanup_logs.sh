#!/bin/bash

# Script de Limpeza de Logs do SOMA
# Uso: ./cleanup_logs.sh [dias] [tamanho_maximo_mb]

LOG_DIR="${1:-.}/logs"
DAYS_TO_KEEP="${2:-30}"      # Manter logs dos últimos 30 dias por padrão
MAX_LOG_SIZE_MB="${3:-1024}" # Máximo 1GB para o arquivo principal
BACKUP_SUFFIX=$(date +%Y%m%d_%H%M%S)

echo "================================================"
echo "LIMPEZA DE LOGS - SOMA"
echo "================================================"
echo "Diretório: $LOG_DIR"
echo "Manter logs de: últimos $DAYS_TO_KEEP dias"
echo "Tamanho máximo: ${MAX_LOG_SIZE_MB}MB"
echo ""

# Verificar se o diretório existe
if [ ! -d "$LOG_DIR" ]; then
    echo "❌ Diretório $LOG_DIR não encontrado!"
    exit 1
fi

# 1. Informações antes da limpeza
echo "📊 ANTES DA LIMPEZA:"
du -sh "$LOG_DIR"
echo ""

# 2. Contar arquivos antigos
OLD_LOGS=$(find "$LOG_DIR" -type f -name "soma_dev_*.log" -mtime +$DAYS_TO_KEEP 2>/dev/null | wc -l)
echo "📋 Encontrados $OLD_LOGS logs antigos (> $DAYS_TO_KEEP dias)"
echo ""

# 3. Remover logs antigos
if [ "$OLD_LOGS" -gt 0 ]; then
    echo "🗑️  Removendo logs antigos..."
    find "$LOG_DIR" -type f -name "soma_dev_*.log" -mtime +$DAYS_TO_KEEP -delete -print
    echo "✅ Logs antigos removidos!"
    echo ""
fi

# 4. Rotacionar arquivo principal se muito grande
if [ -f "$LOG_DIR/soma-run.log" ]; then
    SIZE_BYTES=$(stat -c%s "$LOG_DIR/soma-run.log" 2>/dev/null || stat -f%z "$LOG_DIR/soma-run.log" 2>/dev/null)
    SIZE_MB=$((SIZE_BYTES / 1024 / 1024))

    echo "📝 Tamanho de soma-run.log: ${SIZE_MB}MB"

    if [ "$SIZE_MB" -gt "$MAX_LOG_SIZE_MB" ]; then
        echo "⚠️  Arquivo soma-run.log muito grande!"
        echo "🔄 Rotacionando arquivo..."

        # Backup do arquivo grande
        mv "$LOG_DIR/soma-run.log" "$LOG_DIR/soma-run-${BACKUP_SUFFIX}.log"
        echo "✅ Arquivo rotacionado para: soma-run-${BACKUP_SUFFIX}.log"

        # Comprime se disponível gzip
        if command -v gzip &> /dev/null; then
            echo "📦 Comprimindo arquivo..."
            gzip "$LOG_DIR/soma-run-${BACKUP_SUFFIX}.log"
            echo "✅ Arquivo comprimido!"
        fi
        echo ""
    else
        echo "✅ Tamanho dentro do limite"
        echo ""
    fi
fi

# 5. Limpar arquivos vazios
EMPTY_LOGS=$(find "$LOG_DIR" -type f -size 0 | wc -l)
if [ "$EMPTY_LOGS" -gt 0 ]; then
    echo "🧹 Removendo $EMPTY_LOGS arquivos vazios..."
    find "$LOG_DIR" -type f -size 0 -delete -print
    echo "✅ Arquivos vazios removidos!"
    echo ""
fi

# 6. Informações depois da limpeza
echo "📊 DEPOIS DA LIMPEZA:"
du -sh "$LOG_DIR"
echo ""

# 7. Listar arquivos de log
echo "📋 LOGS DISPONÍVEIS:"
ls -lhS "$LOG_DIR" | grep -E "soma|log" | head -15
echo ""

echo "================================================"
echo "✅ LIMPEZA CONCLUÍDA!"
echo "================================================"
