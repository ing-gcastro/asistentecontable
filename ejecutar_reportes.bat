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
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" indexar_pdfs.py

echo.
echo ==============================================
echo 2. Actualizando Reporte Maestro...
echo ==============================================
cd /d "C:\Users\gcastro\ROBOTS\asistentecontable\Proyecto_facturas_v2"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" reporter.py

echo.
echo ==============================================
echo 3. Actualizando Reporte de Legajos de Pago...
echo ==============================================
cd /d "C:\Users\gcastro\ROBOTS\asistentecontable"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" indexar_legajos_pago.py

echo.
echo ¡Proceso finalizado con éxito!
timeout /t 5 >nul
