@echo off
chcp 65001 >nul
title Reporte de Legajos de Pago
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0"
"%PY%" indexar_legajos_pago.py
if %ERRORLEVEL% EQU 0 (echo Reporte actualizado con exito.) else (echo Hubo un problema al actualizar el reporte.)
if not "%~1"=="--sin-pausa" pause
