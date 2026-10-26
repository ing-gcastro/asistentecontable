@echo off
chcp 65001 > nul
title Paso 2 - Procesar Lote
cd /d "%~dp0"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" 2_procesar_lote.py
pause
