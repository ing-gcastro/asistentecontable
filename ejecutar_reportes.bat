@echo off
chcp 65001 >nul
title Actualizar reportes
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
echo ==============================================
echo 1. Reporte Maestro, sectores y detallado
echo ==============================================
cd /d "%~dp0Proyecto_facturas_v2"
"%PY%" reporter.py
echo.
echo ==============================================
echo 2. Reporte de Legajos de Pago (Tesoreria)
echo ==============================================
cd /d "%~dp0tesoreria_pagos"
"%PY%" indexar_legajos_pago.py
echo.
echo Proceso finalizado.
timeout /t 5 >nul
