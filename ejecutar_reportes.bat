@echo off
:: Script para automatizar la unificación de reportes de facturas

echo Cargando el entorno de Anaconda...
:: Cargar el entorno de Anaconda usando la ruta correcta
call "C:\Users\gcastro\AppData\Local\anaconda3\Scripts\activate.bat" base

echo.
echo ==============================================
echo 1. Actualizando Reporte de Autorizacion...
echo ==============================================
cd /d "C:\Users\gcastro\ROBOTS\asistentecontable\REPORTE DE AUTORIZACION"
python indexar_pdfs.py

echo.
echo ==============================================
echo 2. Actualizando Reporte Maestro...
echo ==============================================
cd /d "C:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2"
python reporter.py

echo.
echo ¡Proceso finalizado con éxito!
timeout /t 5 >nul
