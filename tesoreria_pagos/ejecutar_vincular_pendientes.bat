@echo off
chcp 65001 >nul
title Vinculacion de Mails de Pago y Razon Social
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0"
"%PY%" vincular_avisos_pendientes.py
if not "%~1"=="--sin-pausa" pause
