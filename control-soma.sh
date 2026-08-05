#!/bin/bash

# 🎮 Script de Controle do SOMA via PM2
# Simplifica: stop, start, restart, status, logs, monit

usage() {
    echo "Uso: $0 <comando>"
    echo ""
    echo "Comandos:"
    echo "  status   - Ver status dos processos"
    echo "  start    - Iniciar orquestrador-soma"
    echo "  stop     - Parar orquestrador-soma"
    echo "  restart  - Reiniciar orquestrador-soma"
    echo "  logs     - Ver últimos logs (30 linhas)"
    echo "  logs-f   - Ver logs em tempo real (Ctrl+C para sair)"
    echo "  monit    - Dashboard de monitoramento (Ctrl+C para sair)"
    echo "  health   - Verificar saúde do sistema"
    echo ""
    exit 1
}

if [ $# -eq 0 ]; then
    usage
fi

COMMAND=$1

# Verificar se pm2 está instalado
if ! command -v pm2 &> /dev/null; then
    echo "❌ pm2 não está instalado!"
    echo "Para instalar: npm install -g pm2"
    exit 1
fi

case $COMMAND in
    status)
        echo "📊 Status dos Processos"
        echo "================================================"
        pm2 status
        echo ""
        ;;

    start)
        echo "▶️  Iniciando bot-igreja..."
        pm2 start bot-igreja || {
            echo "❌ Erro ao iniciar"
            exit 1
        }
        sleep 2
        pm2 status
        echo ""
        ;;

    stop)
        echo "⏹️  Parando bot-igreja..."
        pm2 stop bot-igreja || {
            echo "❌ Erro ao parar"
            exit 1
        }
        sleep 1
        pm2 status
        echo ""
        ;;

    restart)
        echo "🔄 Reiniciando bot-igreja..."
        pm2 restart bot-igreja || {
            echo "❌ Erro ao reiniciar"
            exit 1
        }
        sleep 2
        echo "✅ Reiniciado"
        pm2 status
        echo ""
        ;;

    logs)
        echo "📋 Últimos Logs (30 linhas)"
        echo "================================================"
        pm2 logs bot-igreja --lines 30 --nostream
        echo ""
        ;;

    logs-f)
        echo "📋 Logs em Tempo Real (Ctrl+C para sair)"
        echo "================================================"
        pm2 logs bot-igreja
        ;;

    monit)
        echo "📊 Dashboard de Monitoramento (Ctrl+C para sair)"
        echo "================================================"
        pm2 monit
        ;;

    health)
        echo "💚 Verificação de Saúde do Sistema"
        echo "================================================"
        echo ""
        echo "1. Status dos processos:"
        pm2 status
        echo ""
        echo "2. CPU e Memória:"
        pm2 monit --summary 2>/dev/null || {
            ps aux | grep -i bot-igreja | grep -v grep | awk '{print "   CPU: " $3 "% | MEM: " $4 "%"}'
        }
        echo ""
        echo "3. Últimas linhas do log:"
        pm2 logs bot-igreja --lines 10 --nostream
        echo ""
        ;;

    *)
        echo "❌ Comando desconhecido: $COMMAND"
        echo ""
        usage
        ;;
esac
