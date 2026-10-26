@echo off
chcp 65001 > nul
title Paso 2 - Procesar Lote de Facturas
echo ======================================================
echo    PASO 2: PROCESAR Y APLICAR LOTE DE FACTURAS
echo ======================================================
echo.

cd /d "%~dp0Proyecto_facturas_v2"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" 2_procesar_lote.py

echo.
echo Presione cualquier tecla para salir...
pause > nul
