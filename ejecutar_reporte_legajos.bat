@echo off
:: Script para actualizar el Reporte de Legajos de Pago 2026

echo ====================================================
echo    ACTUALIZACION DE REPORTE DE LEGAJOS DE PAGO
echo ====================================================
echo.

set "SCRIPT_PATH=%~dp0indexar_legajos_pago.py"

if exist "C:\Users\%USERNAME%\AppData\Local\anaconda3\python.exe" (
    "C:\Users\%USERNAME%\AppData\Local\anaconda3\python.exe" "%SCRIPT_PATH%"
) else if exist "C:\Users\gcastro\AppData\Local\anaconda3\python.exe" (
    "C:\Users\gcastro\AppData\Local\anaconda3\python.exe" "%SCRIPT_PATH%"
) else (
    python "%SCRIPT_PATH%"
)

echo.
if %ERRORLEVEL% EQU 0 (
    echo Reporte de Legajos de Pago actualizado con exito!
) else (
    echo Hubo un problema al actualizar el reporte.
)

echo.
timeout /t 5
