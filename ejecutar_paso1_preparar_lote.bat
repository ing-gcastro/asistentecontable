@echo off
chcp 65001 > nul
title Paso 1 - Preparar Lote de Facturas
echo ======================================================
echo    PASO 1: PREPARAR LOTE DE FACTURAS DE ENTRADA
echo ======================================================
echo.

cd /d "%~dp0Proyecto_facturas_v2"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" 1_preparar_lote.py

echo.
echo Presione cualquier tecla para salir...
pause > nul
