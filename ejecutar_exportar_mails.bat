@echo off
chcp 65001 >nul
title Exportador de Mails Enviados a Excel
echo ======================================================
echo    EXPORTADOR DE MAILS ENVIADOS DE OUTLOOK A EXCEL
echo ======================================================
echo.

cd /d "%~dp0"
python exportar_mails_enviados.py --abrir

echo.
echo Presione cualquier tecla para salir...
pause >nul
