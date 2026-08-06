@echo off
:: Espera 10 segundos para asegurar que la red y la VPN estén listas
timeout /t 10 /nobreak >nul

:: Cambia a la carpeta de tus scripts y ejecuta Python
cd /d C:\Users\gcastro\ROBOTS\REPORTE DE AUTORIZACION
python indexar_pdfs.py