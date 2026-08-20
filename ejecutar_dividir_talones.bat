@echo off
chcp 65001 > nul
echo ===================================================
echo     DIVISION DE TALONES DE CIRCULACION POR PATENTE
echo ===================================================

if "%~1"=="" (
    set "INPUT_FILE=C:\Users\gcastro\Desktop\Talon_Circulacion_10983577_0.pdf"
) else (
    set "INPUT_FILE=%~1"
)

python "%~dp0dividir_talones_patente.py" "%INPUT_FILE%"

echo.
pause
