@echo off
chcp 65001 >nul
title Paso 2 - Procesar Lote de Facturas
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0Proyecto_facturas_v2"
"%PY%" 2_procesar_lote.py
