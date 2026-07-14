@echo off
setlocal
cd /d "%~dp0"
title Gerar EXE - Preparador de Divulgacoes

echo.
echo ==========================================
echo   PREPARADOR DE DIVULGACOES
echo   Desenvolvido por Deivid Brito
echo ==========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PY=py"
    goto :python_ok
)

where python >nul 2>nul
if %errorlevel%==0 (
    set "PY=python"
    goto :python_ok
)

echo Python nao foi encontrado.
echo Instale o Python e marque "Add Python to PATH".
pause
exit /b 1

:python_ok
echo Instalando bibliotecas...
%PY% -m pip install --upgrade pip
if errorlevel 1 goto :error

%PY% -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 goto :error

echo.
echo Gerando executavel unico...
%PY% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "Preparador_de_Divulgacoes" ^
  --collect-all customtkinter ^
  --collect-submodules win32com ^
  --hidden-import win32timezone ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import win32com.client ^
  "%~dp0main.py"

if errorlevel 1 goto :error

if not exist "%~dp0PARA_ENVIAR" mkdir "%~dp0PARA_ENVIAR"
copy /Y "%~dp0dist\Preparador_de_Divulgacoes.exe" "%~dp0PARA_ENVIAR\Preparador_de_Divulgacoes.exe" >nul

echo.
echo ==========================================
echo EXECUTAVEL GERADO COM SUCESSO
echo ==========================================
echo.
echo Envie somente:
echo PARA_ENVIAR\Preparador_de_Divulgacoes.exe
echo.
explorer "%~dp0PARA_ENVIAR"
pause
exit /b 0

:error
echo.
echo Nao foi possivel gerar o executavel.
echo Tire uma foto desta tela para verificarmos.
pause
exit /b 1
