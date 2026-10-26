@echo off
chcp 65001 > nul
echo ===================================================
echo     DIVISION DE TALONES DE CIRCULACION POR PATENTE
echo ===================================================

if "%~1"=="" (
    set "INPUT_FILE=C:\Users\gcastro\Desktop\Frente_Poliza_11101290_0.pdf"
) else (
    set "INPUT_FILE=%~1"
)

if "%~2"=="" (
    set "OUTPUT_FOLDER=C:\Users\gcastro\ROBOTS\SEGUROS"
) else (
    set "OUTPUT_FOLDER=%~2"
)

"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" "%~dp0dividir_talones_patente.py" "%INPUT_FILE%" "%OUTPUT_FOLDER%"

echo.
pause

