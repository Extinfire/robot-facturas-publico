@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Bot Facturas SAT
echo Iniciando Bot Facturas SAT...
echo.

:: Detectar Python 3.12
py -3.12 --version >nul 2>&1
if not errorlevel 1 goto usar312

python --version >nul 2>&1
if not errorlevel 1 goto usarpython

py --version >nul 2>&1
if not errorlevel 1 goto usarpy

echo ERROR: Python no encontrado.
pause
exit /b 1

:usar312
set PY=py -3.12
goto verificar

:usarpython
set PY=python
goto verificar

:usarpy
set PY=py
goto verificar

:verificar
%PY% --version
echo.

:: Instalar dependencias automaticamente si faltan
%PY% -c "import requests" >nul 2>&1
if errorlevel 1 goto instalar_deps
%PY% -c "import playwright" >nul 2>&1
if errorlevel 1 goto instalar_deps
%PY% -c "import telegram" >nul 2>&1
if errorlevel 1 goto instalar_deps
goto arrancar

:instalar_deps
echo Instalando dependencias faltantes...
%PY% -m pip install -r requirements.txt
echo.
echo Instalando Chromium...
%PY% -m playwright install chromium >nul 2>&1
echo.

:arrancar
echo Presiona Ctrl+C para detener.
echo.
%PY% telegram_bot.py
echo.
echo El bot se detuvo.
pause
