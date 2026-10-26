@echo off
:: Espera 10 segundos para asegurar que la red y la VPN estén listas
timeout /t 10 /nobreak >nul

:: Cambia a la carpeta de tus scripts y ejecuta Python
cd /d "%~dp0"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" indexar_pdfs.py