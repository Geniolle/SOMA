#!/bin/bash

# 🧪 Script para testar e iniciar SOMA no servidor

echo "================================================"
echo "🧪 Teste SOMA - Verificar Chrome e Iniciar"
echo "================================================"
echo ""

# 1. Verificar Chrome
echo "1️⃣ Verificando Chrome..."
CHROME_VERSION=$(google-chrome --version)
if [ $? -eq 0 ]; then
    echo "   ✅ $CHROME_VERSION"
else
    echo "   ❌ Chrome não encontrado!"
    exit 1
fi
echo ""

# 2. Verificar ChromeDriver
echo "2️⃣ Verificando ChromeDriver..."
CHROMEDRIVER_VERSION=$(chromedriver --version 2>/dev/null || chromium-chromedriver --version)
if [ $? -eq 0 ]; then
    echo "   ✅ $CHROMEDRIVER_VERSION"
else
    echo "   ❌ ChromeDriver não encontrado!"
    echo "   Instalando..."
    sudo apt-get install -y chromium-chromedriver
    CHROMEDRIVER_VERSION=$(chromedriver --version)
    echo "   ✅ $CHROMEDRIVER_VERSION"
fi
echo ""

# 3. Verificar PM2
echo "3️⃣ Verificando PM2..."
if command -v pm2 &> /dev/null; then
    echo "   ✅ PM2 encontrado"
else
    echo "   ❌ PM2 não encontrado!"
    exit 1
fi
echo ""

# 4. Registrar SOMA no PM2 (se não estiver)
echo "4️⃣ Registrando SOMA no PM2..."
if pm2 list | grep -q "soma-automation"; then
    echo "   ✅ soma-automation já registrado"
else
    echo "   Registrando novo..."
    pm2 start ecosystem.config.js --name soma-automation
    sleep 2
fi
echo ""

# 5. Iniciar/Reiniciar SOMA
echo "5️⃣ Reiniciando SOMA..."
pm2 restart soma-automation
sleep 2
echo ""

# 6. Status
echo "6️⃣ Status Final:"
pm2 status
echo ""

# 7. Logs (últimas 20 linhas)
echo "7️⃣ Logs SOMA (últimas 20 linhas):"
echo "---"
pm2 logs soma-automation --lines 20 --nostream 2>/dev/null || echo "Nenhum log ainda"
echo "---"
echo ""

echo "================================================"
echo "✅ TESTE COMPLETADO!"
echo "================================================"
echo ""
echo "Comandos úteis:"
echo "   pm2 logs soma-automation         # Ver logs"
echo "   pm2 logs soma-automation logs-f  # Tempo real"
echo "   pm2 monit                        # Dashboard"
echo "   bash control-soma.sh health      # Health check"
echo ""
