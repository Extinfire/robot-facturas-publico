@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Instalacion Bot Facturas SAT
echo ============================================
echo   INSTALACION BOT FACTURAS SAT
echo   Solo espera, todo es automatico.
echo ============================================
echo.

:: =============================================
:: PASO 1: Visual C++ Redistributable
:: =============================================
echo [1/6] Verificando Visual C++ Redistributable...
reg query "HKLM\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64" >nul 2>&1
if not errorlevel 1 goto vc_ok
echo No encontrado. Descargando e instalando...
powershell -Command "Invoke-WebRequest -Uri 'https://aka.ms/vs/17/release/vc_redist.x64.exe' -OutFile '%TEMP%\vc_redist.exe'" >nul 2>&1
if not exist "%TEMP%\vc_redist.exe" goto vc_manual
echo Instalando Visual C++...
"%TEMP%\vc_redist.exe" /install /quiet /norestart
del "%TEMP%\vc_redist.exe" >nul 2>&1
echo OK
echo.
goto paso2

:vc_manual
echo No se pudo descargar. Descargalo manualmente:
echo https://aka.ms/vs/17/release/vc_redist.x64.exe
echo Instalalo y vuelve a ejecutar setup.bat
pause
exit /b 1

:vc_ok
echo OK - Ya instalado
echo.

:: =============================================
:: PASO 2: Python 3.12
:: =============================================
:paso2
echo [2/6] Verificando Python 3.12...

py -3.12 --version >nul 2>&1
if not errorlevel 1 goto python_listo

echo Python 3.12 no encontrado.
echo Intentando instalar con py...
py install 3.12 >nul 2>&1
py -3.12 --version >nul 2>&1
if not errorlevel 1 goto python_listo

echo Descargando Python 3.12...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%TEMP%\python312.exe'" >nul 2>&1
if not exist "%TEMP%\python312.exe" goto error_python
echo Instalando Python 3.12...
"%TEMP%\python312.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1
del "%TEMP%\python312.exe" >nul 2>&1
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312\;%LOCALAPPDATA%\Programs\Python\Python312\Scripts\;%PATH%"
py -3.12 --version >nul 2>&1
if not errorlevel 1 goto python_listo
python --version >nul 2>&1
if not errorlevel 1 goto python_generico
goto error_python

:python_listo
set PY=py -3.12
py -3.12 --version
echo OK
echo.
goto paso3

:python_generico
set PY=python
python --version
echo OK
echo.
goto paso3

:error_python
echo ERROR: No se pudo instalar Python 3.12.
echo Descargalo manualmente: https://www.python.org/downloads/release/python-3127/
echo IMPORTANTE: Marca "Add Python to PATH" al instalar.
pause
exit /b 1

:: =============================================
:: PASO 3: Dependencias pip
:: =============================================
:paso3
echo [3/6] Instalando dependencias de Python...
%PY% -m pip install --upgrade pip >nul 2>&1
%PY% -m pip install -r requirements.txt
if errorlevel 1 goto error_deps
:: Reinstalar greenlet por si hay problema de DLL
%PY% -m pip install --force-reinstall greenlet >nul 2>&1
echo OK
:: Verificar pyzbar (lector de QR de facturas). Ya viene en requirements.txt,
:: pero su DLL nativa a veces no carga en Windows. Es OPCIONAL: si falla, el
:: bot usa cv2 como respaldo y el lector de QR sigue funcionando.
%PY% -c "import pyzbar.pyzbar" >nul 2>&1
if errorlevel 1 (
    echo    Lector QR: pyzbar no cargo - se usara cv2 como respaldo ^(OK^).
) else (
    echo    Lector QR: pyzbar OK.
)
echo.
goto paso4

:error_deps
echo ERROR: Fallo la instalacion de dependencias.
echo Verifica tu conexion a internet.
pause
exit /b 1

:: =============================================
:: PASO 4: Chromium para Playwright
:: =============================================
:paso4
echo [4/6] Instalando navegador Chromium...
%PY% -m playwright install chromium >nul 2>&1
if not errorlevel 1 goto chromium_ok

echo No se pudo descargar Chromium de Playwright.
echo Buscando Chrome o Edge en el sistema...
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" goto chrome_ok
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" goto chrome_ok
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" goto edge_ok
if exist "C:\Program Files\Microsoft\Edge\Application\msedge.exe" goto edge_ok
echo ERROR: No se encontro Chrome ni Edge.
echo Instala Google Chrome y vuelve a ejecutar setup.bat
pause
exit /b 1

:chromium_ok
echo OK - Chromium de Playwright instalado
echo.
goto paso5

:chrome_ok
echo OK - Usara Google Chrome del sistema
echo.
goto paso5

:edge_ok
echo OK - Usara Microsoft Edge del sistema
echo.
goto paso5

:: =============================================
:: PASO 5: Carpetas
:: =============================================
:paso5
echo [5/6] Creando carpetas...
if not exist datos mkdir datos
if not exist downloads mkdir downloads
if not exist data mkdir data
echo OK
echo.

:: =============================================
:: PASO 6: Archivo .env
:: =============================================
echo [6/6] Verificando configuracion...
if exist .env goto env_ok

if exist .env.ejemplo copy .env.ejemplo .env >nul
if exist .env goto env_editar

echo # Telegram> .env
echo TELEGRAM_BOT_TOKEN=TU_TOKEN_AQUI>> .env
echo.>> .env
echo # Anthropic>> .env
echo ANTHROPIC_API_KEY=TU_API_KEY_AQUI>> .env
echo.>> .env
echo # Ollama>> .env
echo OLLAMA_BASE_URL=http://localhost:11434>> .env
echo OLLAMA_MODEL=llama3.2-vision>> .env
echo.>> .env
echo DATA_DIR=./data>> .env
echo DOWNLOAD_DIR=./downloads>> .env
echo.>> .env
echo ADMIN_CHAT_ID=>> .env
echo LOG_LEVEL=INFO>> .env

:env_editar
echo.
echo ==========================================
echo   EDITA el archivo .env con tus tokens
echo ==========================================
notepad .env
goto fin

:env_ok
echo .env ya existe, OK
echo.

:fin
if not exist plantilla_reporte.xlsx echo ADVERTENCIA: Falta plantilla_reporte.xlsx

echo.
echo ============================================
echo   INSTALACION COMPLETADA
echo ============================================
echo.
echo Todo listo. Ejecuta iniciar.bat
echo.
pause
