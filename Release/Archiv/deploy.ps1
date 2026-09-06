<#
.SYNOPSIS
Baut und deployt PC-Konfigurator-GUI vollständig in das lokale AppData des Benutzers.

.DESCRIPTION
Zielstruktur:
  %LOCALAPPDATA%\PC-Konfigurator-GUI\
    PC-Konfigurator-GUI.exe
    Datei-Vorlagen\
    Fonts\
    docs\
    Pin-Desktop-Schnellzugriff.ps1
    README.MD

Die Anwendung verwendet das Verzeichnis der EXE als ScriptRoot. Deshalb liegen
alle zur Laufzeit benötigten Dateien direkt neben der deployten EXE.
#>
[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [string]$DeploymentRoot = (Join-Path $env:LOCALAPPDATA 'PC-Konfigurator-GUI'),
    [string]$DocumentationSource
)

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$defaultDocumentationSource = Join-Path $projectRoot 'docs'
if ([string]::IsNullOrWhiteSpace($DocumentationSource)) {
    $DocumentationSource = $defaultDocumentationSource
}
$sourceScript = Join-Path $projectRoot 'src\PC-Konfigurator-GUI.ps1'
$buildScript = Join-Path $projectRoot 'build.ps1'
$buildExe = Join-Path $projectRoot 'build\PC-Konfigurator-GUI.exe'

function Assert-SourcePath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Benötigte Quelle fehlt: $Path"
    }
}

Assert-SourcePath -Path $sourceScript
Assert-SourcePath -Path (Join-Path $projectRoot 'Datei-Vorlagen')
Assert-SourcePath -Path (Join-Path $projectRoot 'Fonts')
Assert-SourcePath -Path (Join-Path $projectRoot 'Pin-Desktop-Schnellzugriff.ps1')

$running = Get-CimInstance Win32_Process -Filter "Name='PC-Konfigurator-GUI.exe'" -ErrorAction SilentlyContinue
if ($running) {
    $ids = ($running.ProcessId -join ', ')
    throw "PC-Konfigurator-GUI läuft noch (PID $ids). Bitte alle GUI-Fenster schließen und den Deploy erneut starten."
}

if (-not $SkipBuild) {
    Write-Host 'Baue aktuelle GUI-EXE ...' -ForegroundColor Cyan
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $buildScript
    if ($LASTEXITCODE -ne 0) {
        throw "Build fehlgeschlagen (ExitCode $LASTEXITCODE)."
    }
}

Assert-SourcePath -Path $buildExe
if (-not (Test-Path -LiteralPath $DocumentationSource -PathType Container)) {
    throw "Dokumentationsquelle fehlt: $DocumentationSource"
}

New-Item -ItemType Directory -Path $DeploymentRoot -Force | Out-Null
Write-Host "Deploye nach: $DeploymentRoot" -ForegroundColor Cyan

Copy-Item -LiteralPath $buildExe -Destination (Join-Path $DeploymentRoot 'PC-Konfigurator-GUI.exe') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'Datei-Vorlagen') -Destination $DeploymentRoot -Recurse -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'Fonts') -Destination $DeploymentRoot -Recurse -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'Pin-Desktop-Schnellzugriff.ps1') -Destination $DeploymentRoot -Force

$docsTarget = Join-Path $DeploymentRoot 'docs'
New-Item -ItemType Directory -Path $docsTarget -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.MD') -Destination (Join-Path $DeploymentRoot 'README.MD') -Force
foreach ($documentationItem in (Get-ChildItem -LiteralPath $DocumentationSource -Force)) {
    Copy-Item -LiteralPath $documentationItem.FullName -Destination $docsTarget -Recurse -Force
}

$deployedExe = Get-Item (Join-Path $DeploymentRoot 'PC-Konfigurator-GUI.exe')
$deployedTemplateCount = (Get-ChildItem -LiteralPath (Join-Path $DeploymentRoot 'Datei-Vorlagen') -File -Recurse).Count
$deployedFontCount = (Get-ChildItem -LiteralPath (Join-Path $DeploymentRoot 'Fonts') -File -Recurse).Count
$deployedDocCount = (Get-ChildItem -LiteralPath $docsTarget -File -Recurse).Count

Write-Host "Deploy erfolgreich:" -ForegroundColor Green
Write-Host "  EXE: $($deployedExe.FullName) ($([Math]::Round($deployedExe.Length / 1KB, 1)) KB)"
Write-Host "  Datei-Vorlagen: $deployedTemplateCount Dateien"
Write-Host "  Fonts: $deployedFontCount Dateien"
Write-Host "  Dokumentationen: $deployedDocCount Dateien"
Write-Host "  Start: $($deployedExe.FullName)" -ForegroundColor Green
