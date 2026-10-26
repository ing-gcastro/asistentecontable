@echo off
chcp 65001 >nul
title Paso 1 - Preparar Lote de Facturas
call "%~dp0scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0Proyecto_facturas_v2"
"%PY%" 1_preparar_lote.py
