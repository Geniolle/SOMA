#!/bin/bash

# 🚀 Script para instalar Chrome no servidor Ubuntu
# Necessário para que SOMA funcione com Selenium

set -e

echo "================================================"
echo "Instalando Google Chrome + ChromeDriver"
echo "================================================"
echo ""

# 1. Adicionar repositório Chrome
echo "1️⃣ Adicionando repositório do Google Chrome..."
curl -fsSL https://dl-ssl.google.com/linux/linux_signing_key.pub | sudo apt-key add -
sudo sh -c 'echo "deb [arch=amd64] http://dl.google.com/linux/chrome/deb/ stable main" >> /etc/apt/sources.list.d/google-chrome.list'
echo "   ✅ Repositório adicionado"
echo ""

# 2. Atualizar apt
echo "2️⃣ Atualizando apt..."
sudo apt-get update -qq
echo "   ✅ apt atualizado"
echo ""

# 3. Instalar Chrome
echo "3️⃣ Instalando Google Chrome (stable)..."
sudo apt-get install -y google-chrome-stable
CHROME_VERSION=$(google-chrome --version)
echo "   ✅ $CHROME_VERSION instalado"
echo ""

# 4. Instalar ChromeDriver
echo "4️⃣ Instalando ChromeDriver..."
sudo apt-get install -y chromium-chromedriver
echo "   ✅ ChromeDriver instalado"
echo ""

# 5. Verificar instalações
echo "5️⃣ Verificando instalações..."
echo ""
echo "Chrome:"
google-chrome --version
echo ""
echo "ChromeDriver:"
chromedriver --version || chromium-chromedriver --version || true
echo ""

# 6. Resumo
echo "================================================"
echo "✅ CHROME INSTALADO COM SUCESSO!"
echo "================================================"
echo ""
echo "Próximo passo:"
echo "   pm2 restart soma-automation"
echo ""
echo "Ou manualmente:"
echo "   cd ~/soma-automation/SOMA"
echo "   .venv/bin/python main.py"
echo ""
