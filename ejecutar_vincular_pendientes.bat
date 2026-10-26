@echo off
chcp 65001 > nul
title Vinculación de Mails de Pago y Razón Social
echo =======================================================================
echo    VINCULACIÓN: REPORTES DE AVISO DE PAGO PENDIENTES - RAZÓN SOCIAL
echo =======================================================================
echo.

python "%~dp0vincular_avisos_pendientes.py"

echo.
pause
