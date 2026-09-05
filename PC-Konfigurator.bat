@echo off
setlocal

title PC-Konfigurator Starter
cd /d "%~dp0"

set "RELEASE_VERSION=__RELEASE_VERSION__"
set "RELEASE_DATE=__RELEASE_DATE__"

set "SCRIPT_PATH=%~dp0src\PC-Konfigurator.ps1"

if not exist "%SCRIPT_PATH%" (
	echo FEHLER: Skript wurde nicht gefunden:
	echo %SCRIPT_PATH%
	echo.
	echo Bitte pruefen Sie den Installationsordner.
	pause
	exit /b 2
)

echo Starte PC-Konfigurator... Version %RELEASE_VERSION% - erstellt am %RELEASE_DATE%
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
	echo Der PC-Konfigurator wurde mit Fehlercode %EXITCODE% beendet.
	echo Pruefen Sie die Log-Datei unter:
	echo %USERPROFILE%\Documents\PC-Konfigurator\Logs\
) else (
	echo Der PC-Konfigurator wurde erfolgreich beendet.
)

echo.
exit /b %EXITCODE%