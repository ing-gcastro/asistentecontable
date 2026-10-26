@echo off
chcp 65001 >nul
title Exportador de Mails Enviados a Excel
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0"
"%PY%" exportar_mails_enviados.py --abrir
if not "%~1"=="--sin-pausa" pause
