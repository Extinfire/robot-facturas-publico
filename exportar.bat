@echo off
chcp 65001 >nul 2>&1
title Exportar Bot Facturas SAT
echo ============================================
echo   EXPORTAR BOT (sin datos personales)
echo ============================================
echo.

for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
set ZIP_NAME=bot_facturas_sat_LIMPIO_%FECHA%.zip

echo Creando %ZIP_NAME%...
echo.

:: Solo codigo + plantilla + config ejemplo + catalogo empresas
powershell -Command ^
  "Compress-Archive -Force -Path 'telegram_bot.py','sat_invoice_bot.py','sat_verificador.py','requirements.txt','setup.bat','iniciar.bat','plantilla_reporte.xlsx','.env.ejemplo' -DestinationPath '%ZIP_NAME%'"

:: Agregar catalogo de empresas (solo empresas.json, no cache)
if exist data\empresas.json (
    echo Incluyendo catalogo de empresas...
    powershell -Command "Compress-Archive -Update -Path 'data\empresas.json' -DestinationPath '%ZIP_NAME%'"
)

echo.
echo ============================================
echo   ZIP creado: %ZIP_NAME%
echo ============================================
echo.
echo Incluido:
echo   - Codigo del bot
echo   - Catalogo de empresas (NIT)
echo   - Plantilla de reportes
echo   - .env.ejemplo (sin tokens)
echo.
echo NO incluido:
echo   - Tokens ni claves
echo   - Datos de usuarios
echo   - PDFs descargados
echo.
echo El destinatario debe:
echo   1. Instala Python 3.10+ (marca "Add to PATH")
echo   2. Descomprime el ZIP
echo   3. Renombra .env.ejemplo a .env
echo   4. Coloca sus tokens en .env
echo   5. Ejecuta: setup.bat
echo   6. Ejecuta: iniciar.bat
echo.
pause
