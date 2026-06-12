# ===================================================================
# dev-patch.ps1 - Patch local instantaneo (sem GitHub Actions)
# ===================================================================
# Uso:
#   .\dev-patch.ps1            # patch completo (asar + plugin)
#   .\dev-patch.ps1 -Plugin    # so plugin (skip asar, ~2x mais rapido)
#   .\dev-patch.ps1 -NoLaunch  # patch sem reabrir o app
#
# Pra liberar pros outros (apos testar): .\dev-release.ps1
# ===================================================================

param(
    [switch]$Plugin,
    [switch]$NoLaunch,
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"

$src = $PSScriptRoot
$res = "$env:LOCALAPPDATA\Programs\lion-workspace\resources"
$cep = "$env:APPDATA\Adobe\CEP\extensions\com.lionworkspace.premiere"
$tmp = "$env:TEMP\lion-asar-patch"
$asarFile = "$res\app.asar"
$unpackedDir = "$res\app.asar.unpacked"

if (-not (Test-Path $res)) {
    Write-Host "ERRO: Lion Workspace nao instalado em $res" -ForegroundColor Red
    Write-Host "Instale via GitHub release antes." -ForegroundColor Red
    exit 1
}

$t0 = Get-Date

# 1) Mata processo se rodando
$running = Get-Process | Where-Object { $_.Name -eq 'Lion Workspace' }
if ($running) {
    Write-Host "Fechando Lion Workspace..." -ForegroundColor Yellow
    $running | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 500
}

# 2) Plugin Premiere (CEP)
if (Test-Path $cep) {
    Copy-Item "$src\premiere-plugin\client\index.html" "$cep\client\index.html" -Force
    Copy-Item "$src\premiere-plugin\host\index.jsx" "$cep\host\index.jsx" -Force
    if (Test-Path "$res\premiere-plugin") {
        Copy-Item "$src\premiere-plugin\client\index.html" "$res\premiere-plugin\client\index.html" -Force
        Copy-Item "$src\premiere-plugin\host\index.jsx" "$res\premiere-plugin\host\index.jsx" -Force
    }
    Write-Host "[OK] Plugin CEP atualizado" -ForegroundColor Green
} else {
    Write-Host "[!] Plugin CEP nao instalado em $cep - skip" -ForegroundColor Yellow
}

if ($Plugin) {
    $elapsed = (Get-Date) - $t0
    Write-Host ""
    Write-Host "PLUGIN PATCHED em $($elapsed.TotalSeconds.ToString('0.0'))s" -ForegroundColor Cyan
    Write-Host "Reinicie o Premiere pra plugin recarregar."
    if (-not $NoLaunch -and (Test-Path "$res\..\Lion Workspace.exe")) {
        Start-Process -FilePath "$res\..\Lion Workspace.exe"
    }
    exit 0
}

# 3) Asar repack
$htmlFiles = @('index.html', 'lion-search.html', 'rotoscope-editor.html', 'mask-editor.html', 'bg-video-worker.html')
foreach ($html in $htmlFiles) {
    if (Test-Path "$src\$html") {
        Copy-Item "$src\$html" "$unpackedDir\$html" -Force
        if ($Verbose) { Write-Host "  $html -> unpacked" -ForegroundColor DarkGray }
    }
}

if (-not (Test-Path "$tmp\main.js")) {
    Write-Host "Extraindo asar inicial..." -ForegroundColor Yellow
    if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
    New-Item -ItemType Directory -Path $tmp | Out-Null
    & cmd /c "npx --yes @electron/asar extract `"$asarFile`" `"$tmp`" 2>&1" | Out-Null
}

$jsFiles = @('main.js', 'preload.js')
foreach ($f in $jsFiles) {
    if (Test-Path "$src\$f") {
        Copy-Item "$src\$f" "$tmp\$f" -Force
    }
}
foreach ($html in $htmlFiles) {
    if (Test-Path "$src\$html") { Copy-Item "$src\$html" "$tmp\$html" -Force }
}

Write-Host "Repacking asar..." -ForegroundColor Yellow
$unpackPattern = 'node_modules/{@huggingface,@imgly,@mediapipe,@img,sharp,onnxruntime-node,onnxruntime-common,onnxruntime-web,@xenova}'
& cmd /c "npx --yes @electron/asar pack `"$tmp`" `"$asarFile`" --unpack-dir `"$unpackPattern`" --unpack `"*.html`" 2>&1" | Out-Null

if (-not $NoLaunch -and (Test-Path "$res\..\Lion Workspace.exe")) {
    Start-Process -FilePath "$res\..\Lion Workspace.exe"
}

$elapsed = (Get-Date) - $t0
Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  PATCHED LOCAL em $($elapsed.TotalSeconds.ToString('0.0'))s" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "App relancado. Plugin atualizado."
Write-Host "Pra plugin: feche e reabra o Premiere."
Write-Host "Pra liberar pros outros: .\dev-release.ps1"
