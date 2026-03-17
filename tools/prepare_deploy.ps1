param(
  [switch]$Force,
  [string]$CredPath
)

$ErrorActionPreference = 'Stop'

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Write-Ok($msg)   { Write-Host "[OK]   $msg" -ForegroundColor Green }

# Root do projeto = pasta atual
$root = (Get-Location).Path

$deployDir = Join-Path $root "deploy"
$deploySecretsDir = Join-Path $deployDir "secrets"

# 1) Pastas
New-Item -ItemType Directory -Force -Path $deployDir | Out-Null
New-Item -ItemType Directory -Force -Path $deploySecretsDir | Out-Null
Write-Ok "Pastas criadas: deploy\ e deploy\secrets\"

# 2) config.env (copiar .env se existir; senão gerar template)
$srcEnv = Join-Path $root ".env"
$dstEnv = Join-Path $deployDir "config.env"

if ((Test-Path $dstEnv) -and (-not $Force)) {
  Write-Warn "deploy\config.env já existe. Use -Force para sobrescrever."
} else {
  if (Test-Path $srcEnv) {
    Copy-Item -Force $srcEnv $dstEnv
    Write-Ok "deploy\config.env criado por cópia do .env"
  } else {
    @"
PYTHONPATH=.\src
HEADLESS=true
ALLOW_RETRY_ERROR=true

# --- Google / Sheets ---
GOOGLE_CREDENTIALS_PATH=.\deploy\secrets\credenciais.json
SPREADSHEET_URL=""
SHEET_CONTAORDEM=TESTE_CONTAORDEM
SHEET_CAIXAS=GERENCIAR CAIXAS

# --- Site (SOMA) ---
SITE_USER=""
SITE_PASSWORD=""
SITE_LOGIN_URL=https://verbodavida.info/apps/index.php
SITE_HOME_URL=https://verbodavida.info/IVV/

# --- Execução ---
TIMEOUT_SECONDS=20
RETRY_COUNT=2
BATCH_SIZE=20

# --- Logs/artefactos ---
LOG_LEVEL=INFO
LOG_DIR=logs
SCREENSHOTS_DIR=artifacts/screenshots
"@ | Set-Content -Encoding UTF8 -Path $dstEnv

    Write-Ok "deploy\config.env criado (template)."
    Write-Warn "Preenche SPREADSHEET_URL, SITE_USER e SITE_PASSWORD (e o que mais precisares)."
  }
}

# 3) run_soma.bat
$batPath = Join-Path $deployDir "run_soma.bat"
if ((Test-Path $batPath) -and (-not $Force)) {
  Write-Warn "deploy\run_soma.bat já existe. Use -Force para sobrescrever."
} else {
  @"
@echo off
setlocal
cd /d %~dp0..\

REM aponta para o env de deploy
set ENV_FILE=.\deploy\config.env

REM executa o módulo (modo dev). Depois, quando gerar EXE, apontas para dist\...\soma_run.exe
.\.venv\Scripts\python.exe -m soma_app.workflows.run_soma

endlocal
"@ | Set-Content -Encoding UTF8 -Path $batPath

  Write-Ok "deploy\run_soma.bat criado."
}

# 4) Copiar credenciais JSON para deploy\secrets\credenciais.json
$dstCred = Join-Path $deploySecretsDir "credenciais.json"

# ordem de descoberta do json:
# 1) parâmetro -CredPath
# 2) env GOOGLE_CREDENTIALS_PATH
# 3) .\secrets\credenciais.json
# 4) .\secrets\credentials.json
$candidate = $null

if ($CredPath -and (Test-Path $CredPath)) {
  $candidate = (Resolve-Path $CredPath).Path
} elseif ($env:GOOGLE_CREDENTIALS_PATH -and (Test-Path $env:GOOGLE_CREDENTIALS_PATH)) {
  $candidate = (Resolve-Path $env:GOOGLE_CREDENTIALS_PATH).Path
} else {
  $c1 = Join-Path $root "credenciais.json"
  $c2 = Join-Path $root "secrets\credenciais.json"
  if (Test-Path $c1) { $candidate = (Resolve-Path $c1).Path }
  elseif (Test-Path $c2) { $candidate = (Resolve-Path $c2).Path }
}

if ($candidate) {
  if ((Test-Path $dstCred) -and (-not $Force)) {
    Write-Warn "deploy\secrets\credenciais.json já existe. Use -Force para sobrescrever."
  } else {
    Copy-Item -Force $candidate $dstCred
    Write-Ok "Credenciais copiadas para deploy\secrets\credenciais.json"
  }
} else {
  Write-Warn "Não encontrei o JSON de credenciais. Copia manualmente para deploy\secrets\credenciais.json"
  Write-Warn "Ou executa: powershell -ExecutionPolicy Bypass -File .\tools\prepare_deploy.ps1 -CredPath 'C:\...\credenciais.json' -Force"
}

# 5) .gitignore (evitar commitar credenciais)
$gitignore = Join-Path $root ".gitignore"
$ignoreLine = "deploy/secrets/"
if (Test-Path $gitignore) {
  $content = Get-Content $gitignore -ErrorAction SilentlyContinue
  if ($content -notcontains $ignoreLine) {
    Add-Content -Encoding UTF8 -Path $gitignore -Value "`n$ignoreLine"
    Write-Ok ".gitignore atualizado: deploy/secrets/"
  } else {
    Write-Info ".gitignore já contém deploy/secrets/"
  }
} else {
  Set-Content -Encoding UTF8 -Path $gitignore -Value "$ignoreLine`n"
  Write-Ok ".gitignore criado com deploy/secrets/"
}

Write-Host ""
Write-Ok "Preparação concluída."
Write-Info "Próximo passo: validar deploy\config.env e testar deploy\run_soma.bat."
