@echo off
echo Atualizando o site Encontro de Contas Enpavi com a planilha atual...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\vigiar_planilha.ps1" -UmaVez -Forcar
echo.
pause
