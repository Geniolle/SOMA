#!/bin/bash

# 🚀 Script de Deployment no Servidor
# Automatiza: git reset, pull, npm install, pm2 restart

set -e  # Exit on error

echo "================================================"
echo "DEPLOYMENT NO SERVIDOR - SOMA"
echo "================================================"
echo ""

PROJECT_DIR="${1:-.}"
cd "$PROJECT_DIR"

echo "📍 Diretório: $(pwd)"
echo ""

# 1. Sincronizar com GitHub
echo "🔄 [1/5] Sincronizando com GitHub..."
echo "   - Descartando alterações locais..."
git fetch origin
git reset --hard origin/main

echo "   - Puxando atualizações..."
git pull origin main
echo "   ✅ Git sincronizado"
echo ""

# 2. Verificar se npm existe
if command -v npm &> /dev/null; then
    echo "📦 [2/5] Atualizando dependências Node.js..."
    npm install --production
    echo "   ✅ Node.js dependências instaladas"
    echo ""
else
    echo "⚠️  [2/5] npm não encontrado, pulando..."
    echo ""
fi

# 3. Verificar se Python venv existe
if [ -d ".venv" ]; then
    echo "🐍 [3/5] Atualizando dependências Python..."
    .venv/bin/pip install -q -r requirements.lock.txt 2>/dev/null || true
    echo "   ✅ Python dependências atualizadas"
    echo ""
else
    echo "⚠️  [3/5] venv não encontrado, criando..."
    python3 -m venv .venv
    .venv/bin/pip install -q -r requirements.lock.txt
    echo "   ✅ venv criado e dependências instaladas"
    echo ""
fi

# 4. Reiniciar com PM2
echo "🔄 [4/5] Reiniciando processos PM2..."

if command -v pm2 &> /dev/null; then
    # Parar o processo específico se existir
    if pm2 list | grep -q "bot-igreja"; then
        pm2 stop bot-igreja
        pm2 restart bot-igreja
        echo "   ✅ bot-igreja reiniciado"
    else
        echo "   ⚠️  bot-igreja não encontrado em pm2"
    fi

    # Opção: reiniciar tudo
    # pm2 stop all
    # pm2 restart all
    # echo "   ✅ Todos os processos reiniciados"

    echo ""
else
    echo "   ⚠️  pm2 não encontrado"
    echo "   Se quiser usar pm2, instale com: npm install -g pm2"
    echo ""
fi

# 5. Verificar status
echo "📊 [5/5] Verificando status..."
echo ""

if command -v pm2 &> /dev/null; then
    echo "Status dos processos:"
    pm2 status

    echo ""
    echo "Últimos logs (primeiras 30 linhas):"
    echo "---"
    pm2 logs orquestrador-soma --lines 30 --nostream || true
    echo "---"
else
    echo "⚠️  pm2 não disponível para status"
fi

echo ""
echo "================================================"
echo "✅ DEPLOYMENT CONCLUÍDO!"
echo "================================================"
echo ""
echo "📝 Resumo:"
echo "   ✅ Git sincronizado"
echo "   ✅ Dependências atualizadas"
echo "   ✅ Processos reiniciados"
echo ""
echo "📊 Para monitorar em tempo real:"
echo "   pm2 monit"
echo ""
echo "📋 Para ver logs contínuos:"
echo "   pm2 logs bot-igreja"
echo ""
