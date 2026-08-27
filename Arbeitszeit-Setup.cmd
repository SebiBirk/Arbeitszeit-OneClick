@echo off
setlocal
title Arbeitszeit Setup
cd /d "%~dp0"

echo.
echo Arbeitszeit wird installiert...
echo.

powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File "%~dp0Install-Arbeitszeit-OneClick.ps1"
set "exitCode=%ERRORLEVEL%"

if not "%exitCode%"=="0" (
    echo.
    echo Installation fehlgeschlagen. Bitte diese Ausgabe an die IT weitergeben.
    echo.
    pause
)

exit /b %exitCode%
