@echo off
chcp 65001 >nul
title Actualizar reportes de autorizacion
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
:: Espera 10 segundos para asegurar que la red y la VPN esten listas
timeout /t 10 /nobreak >nul
cd /d "%~dp0..\Proyecto_facturas_v2"
"%PY%" reporter.py
