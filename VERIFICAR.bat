@echo off
chcp 65001 >nul
title Verificar instalacion - Asistente Contable
cd /d "%~dp0"
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
"%PY%" "%~dp0Proyecto_facturas_v2\verificar_entorno.py"
set "_ENTORNO=%ERRORLEVEL%"
echo.
echo Corriendo tests...
"%PY%" -W ignore -m unittest discover -s "%~dp0Proyecto_facturas_v2\tests" -q
set "_TESTS=%ERRORLEVEL%"
if not "%_TESTS%"=="0" (echo [ERROR] Fallaron tests.) else (echo TESTS OK)
if not "%~1"=="--sin-pausa" pause
if not "%_TESTS%"=="0" exit /b 1
exit /b %_ENTORNO%
