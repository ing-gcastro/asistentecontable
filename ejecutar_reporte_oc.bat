@echo off
chcp 65001 > nul
title Reporte de Ordenes de Compra TELCOM - Antina
echo ===================================================
echo   Procesador de Ordenes de Compra - Antina TELCOM
echo ===================================================
echo.
echo Iniciando el escaneo y extraccion de comprobantes...
python procesar_oc_telcom.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================
    echo Proceso finalizado con exito!
    echo Abriendo reporte_oc_telcom.html en el navegador...
    echo ===================================================
    start reporte_oc_telcom.html
) else (
    echo.
    echo Ocurrio un error al procesar las Ordenes de Compra.
)

pause
