# 🚀 Script de Deploy Local
# Ativa venv Python e roda o orquestrador

param(
    [switch]$Monitor,
    [switch]$Stop,
    [switch]$Status
)

$PROJECT_DIR = "C:\workspace\SOMA"

Write-Host ""
Write-Host "================================================"
Write-Host "DEPLOY LOCAL - SOMA"
Write-Host "================================================"
Write-Host ""

# Verificar se estamos no diretório correto
if (-not (Test-Path "$PROJECT_DIR\.venv")) {
    Write-Host "❌ Ambiente virtual não encontrado em: $PROJECT_DIR"
    Write-Host ""
    Write-Host "Para criar o ambiente virtual, execute:"
    Write-Host "   cd $PROJECT_DIR"
    Write-Host "   python -m venv .venv"
    Write-Host "   .\.venv\Scripts\Activate.ps1"
    Write-Host "   pip install -r requirements.lock.txt"
    exit 1
}

# Ativar venv
Write-Host "🐍 [1/4] Ativando ambiente virtual..."
& "$PROJECT_DIR\.venv\Scripts\Activate.ps1"
Write-Host "   ✅ venv ativado"
Write-Host ""

# Se --Stop
if ($Stop) {
    Write-Host "⏹️  [2/4] Parando orquestrador..."
    if (Get-Command pm2 -ErrorAction SilentlyContinue) {
        pm2 stop orquestrador-soma
        Write-Host "   ✅ Parado"
    } else {
        Write-Host "   ⚠️  pm2 não disponível"
    }
    Write-Host ""
    exit 0
}

# Se --Status
if ($Status) {
    Write-Host "📊 [2/4] Status do PM2..."
    if (Get-Command pm2 -ErrorAction SilentlyContinue) {
        pm2 status
    } else {
        Write-Host "   ⚠️  pm2 não disponível"
    }
    Write-Host ""
    exit 0
}

# Se --Monitor
if ($Monitor) {
    Write-Host "📊 [2/4] Dashboard de Monitoramento..."
    if (Get-Command pm2 -ErrorAction SilentlyContinue) {
        pm2 monit
    } else {
        Write-Host "   ⚠️  pm2 não disponível"
    }
    Write-Host ""
    exit 0
}

# Caso padrão: rodar orquestrador
Write-Host "🚀 [2/4] Iniciando orquestrador..."
Write-Host "   📝 Verificando dependências..."
pip install -q -r requirements.lock.txt
Write-Host "   ✅ Dependências OK"
Write-Host ""

Write-Host "▶️  [3/4] Executando python main.py..."
Write-Host ""
Write-Host "================================================"
Write-Host ""

# Executar
python main.py

Write-Host ""
Write-Host "================================================"
Write-Host "❌ Orquestrador parado"
Write-Host "================================================"
