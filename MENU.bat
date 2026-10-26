@echo off
chcp 65001 >nul
title Asistente Contable - Menu
cd /d "%~dp0"
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)

:menu
cls
echo ======================================================
echo           ASISTENTE CONTABLE - CICLO DE COMPRAS
echo ======================================================
echo   1. Descargar facturas de Outlook
echo   2. PASO 1 - Preparar lote (genera Excel de validacion)
echo   3. PASO 2 - Procesar lote (aplica el Excel)
echo   4. Regenerar todos los reportes
echo   5. Monitor de \Compras (dejar la ventana abierta)
echo   6. Recordatorio de pendientes por sector (mail)
echo   7. Verificar instalacion (tests)
echo   8. Instalar / actualizar dependencias
echo   0. Salir
echo ======================================================
set "OPC="
set /p "OPC=Elegi una opcion: "

if "%OPC%"=="1" ( pushd Proyecto_facturas_v2 & "%PY%" auto_downloader.py & popd & pause & goto menu )
if "%OPC%"=="2" ( pushd Proyecto_facturas_v2 & "%PY%" 1_preparar_lote.py & popd & goto menu )
if "%OPC%"=="3" ( pushd Proyecto_facturas_v2 & "%PY%" 2_procesar_lote.py & popd & goto menu )
if "%OPC%"=="4" ( pushd Proyecto_facturas_v2 & "%PY%" reporter.py & popd & pause & goto menu )
if "%OPC%"=="5" ( pushd Proyecto_facturas_v2 & "%PY%" monitor_compras.py & popd & pause & goto menu )
if "%OPC%"=="6" ( pushd Proyecto_facturas_v2 & "%PY%" notificador_sectores.py --pendientes & popd & pause & goto menu )
if "%OPC%"=="7" ( call "%~dp0VERIFICAR.bat" & goto menu )
if "%OPC%"=="8" ( call "%~dp0INSTALAR.bat" & goto menu )
if "%OPC%"=="0" exit /b 0
goto menu
