@echo off
chcp 65001 >nul 2>&1
title Empaquetar Bot para Migracion
echo ============================================
echo   EMPAQUETANDO BOT PARA MIGRACION
echo ============================================
echo.

:: Nombre del ZIP con fecha
for /f "tokens=1-3 delims=/" %%a in ('date /t') do set FECHA=%%c-%%b-%%a
set ZIP_NAME=bot_facturas_sat_%FECHA%.zip

echo Creando %ZIP_NAME% con TODO incluido...
echo.

:: Crear ZIP con archivos principales + .env real + datos de usuarios + catalogo
powershell -Command ^
  "Compress-Archive -Force -Path 'telegram_bot.py','sat_invoice_bot.py','sat_verificador.py','requirements.txt','setup.bat','iniciar.bat','plantilla_reporte.xlsx','.env','.env.ejemplo' -DestinationPath '%ZIP_NAME%'"

:: Agregar carpeta data (catalogo de empresas)
if exist data (
    echo Incluyendo catalogo de empresas...
    powershell -Command "Compress-Archive -Update -Path 'data' -DestinationPath '%ZIP_NAME%'"
)

:: Agregar carpeta datos (registros y PDFs de usuarios)
if exist datos (
    echo Incluyendo datos de usuarios y PDFs...
    powershell -Command "Compress-Archive -Update -Path 'datos' -DestinationPath '%ZIP_NAME%'"
)

echo.
echo ============================================
echo   ZIP creado: %ZIP_NAME%
echo ============================================
echo.
echo Contenido incluido:
echo   - Codigo del bot
echo   - Tu .env con tokens reales
echo   - Catalogo de empresas
echo   - Datos y PDFs de todos los usuarios
echo   - Plantilla de reportes
echo.
echo En la otra computadora:
echo   1. Instala Python 3.10+ (marca "Add to PATH")
echo   2. Descomprime el ZIP en una carpeta
echo   3. Ejecuta: setup.bat
echo   4. Ejecuta: iniciar.bat
echo.
echo IMPORTANTE: Apaga el bot en este equipo antes
echo de encenderlo en el otro.
echo.
pause
