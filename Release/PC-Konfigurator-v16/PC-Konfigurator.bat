@echo off
setlocal

title PC-Konfigurator Starter
cd /d "%~dp0"

set "SCRIPT_PATH=%~dp0src\PC-Konfigurator.ps1"

if not exist "%SCRIPT_PATH%" (
	echo FEHLER: Skript wurde nicht gefunden:
	echo %SCRIPT_PATH%
	echo.
	echo Bitte pruefen Sie den Installationsordner.
	pause
	exit /b 2
)

echo Starte PC-Konfigurator...
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
pause
exit /b %EXITCODE%