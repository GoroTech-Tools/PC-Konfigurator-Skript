<#
.SYNOPSIS
Installiert ein PC-Konfigurator-GUI-Release in das lokale AppData.
#>
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'PC-Konfigurator-GUI'),
    [switch]$CreateDesktopShortcut
)

$ErrorActionPreference = 'Stop'
$sourceRoot = Split-Path $PSScriptRoot -Parent
$exeSource = Join-Path $sourceRoot 'PC-Konfigurator-GUI.exe'
if (-not (Test-Path -LiteralPath $exeSource -PathType Leaf)) {
    throw "Release-EXE fehlt: $exeSource"
}

$running = Get-CimInstance Win32_Process -Filter "Name='PC-Konfigurator-GUI.exe'" -ErrorAction SilentlyContinue
if ($running) {
    throw "PC-Konfigurator-GUI läuft noch. Bitte zuerst alle GUI-Fenster schließen."
}

New-Item -ItemType Directory -Path $InstallRoot -Force | Out-Null
foreach ($itemName in @('PC-Konfigurator-GUI.exe', 'Datei-Vorlagen', 'Fonts', 'docs', 'Pin-Desktop-Schnellzugriff.ps1', 'README.MD')) {
    $source = Join-Path $sourceRoot $itemName
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Release-Datei fehlt: $source"
    }
    Copy-Item -LiteralPath $source -Destination $InstallRoot -Recurse -Force
}

if ($CreateDesktopShortcut) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $shortcutPath = Join-Path $desktop 'PC-Konfigurator-GUI.lnk'
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = Join-Path $InstallRoot 'PC-Konfigurator-GUI.exe'
    $shortcut.WorkingDirectory = $InstallRoot
    $shortcut.Description = 'PC-Konfigurator-GUI'
    $shortcut.Save()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
}

Write-Host "Installation erfolgreich: $InstallRoot" -ForegroundColor Green
Write-Host "Start: $(Join-Path $InstallRoot 'PC-Konfigurator-GUI.exe')"
