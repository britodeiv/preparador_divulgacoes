@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0PARA_ENVIAR\Preparador_de_Divulgacoes.exe" (
    echo Execute primeiro: 1_GERAR_EXECUTAVEL.bat
    pause
    exit /b 1
)

start "" "%~dp0PARA_ENVIAR\Preparador_de_Divulgacoes.exe"
