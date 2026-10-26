@echo off
chcp 65001 >nul
title Division de talones de circulacion por patente
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
if "%~1"=="" (set "INPUT_FILE=%USERPROFILE%\Desktop\Frente_Poliza_11101290_0.pdf") else (set "INPUT_FILE=%~1")
if "%~2"=="" (set "OUTPUT_FOLDER=%USERPROFILE%\ROBOTS\SEGUROS") else (set "OUTPUT_FOLDER=%~2")
"%PY%" "%~dp0dividir_talones_patente.py" "%INPUT_FILE%" "%OUTPUT_FOLDER%"
if not "%~1"=="--sin-pausa" pause
