@echo off
chcp 65001 >nul
title Monitor de Compras (MEJ-05) - dejar abierto
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0Proyecto_facturas_v2"
"%PY%" monitor_compras.py
if not "%~1"=="--sin-pausa" pause
