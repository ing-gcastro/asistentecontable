@echo off
chcp 65001 >nul
title Descarga de facturas de Outlook
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0Proyecto_facturas_v2"
"%PY%" auto_downloader.py %*
if not "%~1"=="--sin-pausa" pause
