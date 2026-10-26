@echo off
chcp 65001 >nul
title Reporte de Ordenes de Compra TELCOM
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0"
echo Iniciando el escaneo y extraccion de comprobantes...
"%PY%" procesar_oc_telcom.py
if %ERRORLEVEL% EQU 0 (
    echo Proceso finalizado con exito. Abriendo reporte...
    start "" "%~dp0reporte_oc_telcom.html"
) else (
    echo Ocurrio un error al procesar las Ordenes de Compra.
)
if not "%~1"=="--sin-pausa" pause
