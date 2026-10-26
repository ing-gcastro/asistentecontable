@echo off
chcp 65001 > nul
title Paso 1 - Preparar Lote
cd /d "%~dp0"
"C:\Users\gcastro\AppData\Local\anaconda3\python.exe" 1_preparar_lote.py
pause
