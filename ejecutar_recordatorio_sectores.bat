@echo off
chcp 65001 >nul
title Recordatorio de facturas pendientes por sector
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0Proyecto_facturas_v2"
"%PY%" notificador_sectores.py --pendientes
if not "%~1"=="--sin-pausa" pause
