<#
.SYNOPSIS
Erzeugt ein vollständiges, auf andere Rechner ausrollbares GUI-Release-ZIP.
#>
[CmdletBinding()]
param(
    [string]$Version = '1.0.0.0',
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$releaseRoot = Join-Path $projectRoot 'release'
$stageRoot = Join-Path $releaseRoot "PC-Konfigurator-GUI-v$Version"
$zipPath = Join-Path $releaseRoot "PC-Konfigurator-GUI-v$Version.zip"
$buildScript = Join-Path $projectRoot 'build.ps1'
$buildExe = Join-Path $projectRoot 'build\PC-Konfigurator-GUI.exe'

if (Get-CimInstance Win32_Process -Filter "Name='PC-Konfigurator-GUI.exe'" -ErrorAction SilentlyContinue) {
    throw 'PC-Konfigurator-GUI läuft noch. Bitte vor dem Release alle GUI-Prozesse schließen.'
}

if (-not $SkipBuild) {
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $buildScript -Version $Version
    if ($LASTEXITCODE -ne 0) { throw "Build fehlgeschlagen: $LASTEXITCODE" }
}

foreach ($required in @($buildExe, (Join-Path $projectRoot 'Datei-Vorlagen'), (Join-Path $projectRoot 'Fonts'), (Join-Path $projectRoot 'docs'), (Join-Path $projectRoot 'Install-PC-Konfigurator-GUI.ps1'))) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Release-Quelle fehlt: $required" }
}

if (Test-Path $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force }
New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null
Copy-Item -LiteralPath $buildExe -Destination (Join-Path $stageRoot 'PC-Konfigurator-GUI.exe') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.MD') -Destination (Join-Path $stageRoot 'README.MD') -Force

New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path $stageRoot -DestinationPath $zipPath -CompressionLevel Optimal -Force

$files = Get-ChildItem -LiteralPath $stageRoot -File -Recurse
Write-Host "Release erstellt: $zipPath" -ForegroundColor Green
Write-Host "Dateien: $($files.Count); Größe: $([Math]::Round((Get-Item $zipPath).Length / 1MB, 2)) MB"
