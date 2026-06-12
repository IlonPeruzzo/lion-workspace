# ===================================================================
# dev-release.ps1 - Commit + bump versao + push (dispara build cloud)
# Quando tudo testado localmente, libera Win + Mac via auto-update
# ===================================================================
# Uso:
#   .\dev-release.ps1 "mensagem do commit"
#   .\dev-release.ps1                          # pede msg interativo
# ===================================================================

param([string]$Message = "")

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$status = git status --porcelain 2>&1
if (-not $status) {
    Write-Host "Nenhuma mudanca detectada. Nada a commitar." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "=== Mudancas ===" -ForegroundColor Cyan
git status --short
Write-Host ""

if (-not $Message) {
    Write-Host "Mensagem do commit (ou Enter pra cancelar):" -ForegroundColor Yellow
    $Message = Read-Host "  > "
    if (-not $Message) { Write-Host "Cancelado." -ForegroundColor Red; exit 1 }
}

Write-Host ""
Write-Host "Commitando..." -ForegroundColor Yellow
git add -A
git commit -m $Message
if ($LASTEXITCODE -ne 0) { Write-Host "Commit falhou." -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "Bumpando versao..." -ForegroundColor Yellow
$newVer = (npm version patch 2>&1 | Select-Object -Last 1)
if ($LASTEXITCODE -ne 0) { Write-Host "npm version falhou." -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "Pushing..." -ForegroundColor Yellow
git push 2>&1 | Out-Null
git push --tags 2>&1 | Out-Null

Write-Host ""
Write-Host "===============================================" -ForegroundColor Green
Write-Host "  RELEASE $newVer DISPARADO" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Build: https://github.com/IlonPeruzzo/lion-workspace/actions"
Write-Host "  Quando os 2 jobs (Win + Mac) ficarem verdes,"
Write-Host "  usuarios recebem auto-update na proxima abertura."
Write-Host ""
