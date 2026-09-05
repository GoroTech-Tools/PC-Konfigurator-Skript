@echo off
REM filepath: d:\OneDrive\Git-Projekte\PC-Konfigurator\Deploy.bat
title PC-Konfigurator Deployment
cls

echo.
echo ========================================
echo    PC-Konfigurator Deployment
echo ========================================
echo.

REM Prüfe ob PowerShell verfügbar ist
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo FEHLER: PowerShell wurde nicht gefunden!
    echo Bitte stellen Sie sicher, dass PowerShell installiert ist.
    pause
    exit /b 1
)

REM Prüfe ob das Deployment-Skript existiert
if not exist "src\PC-Konfigurator_Deploy.ps1" (
    echo FEHLER: PC-Konfigurator_Deploy.ps1 wurde nicht gefunden!
    echo Stellen Sie sicher, dass die Batch-Datei im gleichen Verzeichnis liegt.
    pause
    exit /b 1
)

echo Starte PC-Konfigurator Deployment...
echo.

REM Starte PowerShell-Skript mit ExecutionPolicy Bypass
powershell.exe -ExecutionPolicy Bypass -File "src\PC-Konfigurator_Deploy.ps1"

REM Prüfe Rückgabewert
if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo    Deployment erfolgreich abgeschlossen!
    echo ========================================
    echo.
    echo Das ZIP-Archiv wurde erstellt.
    echo Pruefen Sie den Release-Ordner.
) else (
    echo.
    echo ========================================
    echo    FEHLER beim Deployment!
    echo ========================================
    echo.
    echo Fehlercode: %errorlevel%
    echo Pruefen Sie die Ausgabe oben fuer Details.
)

echo.