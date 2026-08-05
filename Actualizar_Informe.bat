@echo off
setlocal
set "PATH=%PATH%;C:\Program Files\GitHub CLI"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0actualizar_informe.ps1"
echo.
pause
