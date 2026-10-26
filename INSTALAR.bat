@echo off
chcp 65001 >nul
title Instalar dependencias - Asistente Contable
cd /d "%~dp0"
set "_PAUSA=1"
if /i "%~1"=="--sin-pausa" set "_PAUSA="
set "_COD=0"
echo ======================================================
echo    INSTALACION / ACTUALIZACION DE DEPENDENCIAS
echo ======================================================
echo.

:: Python base (sin .venv) para crear el entorno
set "BASEPY="
if exist "%USERPROFILE%\AppData\Local\anaconda3\python.exe" set "BASEPY=%USERPROFILE%\AppData\Local\anaconda3\python.exe"
if not defined BASEPY if exist "%USERPROFILE%\anaconda3\python.exe" set "BASEPY=%USERPROFILE%\anaconda3\python.exe"
if not defined BASEPY if exist "%ProgramData%\anaconda3\python.exe" set "BASEPY=%ProgramData%\anaconda3\python.exe"
if not defined BASEPY if exist "%USERPROFILE%\miniconda3\python.exe" set "BASEPY=%USERPROFILE%\miniconda3\python.exe"
if not defined BASEPY (
    py -3 -c "import sys" >nul 2>&1 && for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)"') do set "BASEPY=%%P"
)
if not defined BASEPY (
    python -c "import sys" >nul 2>&1 && for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)"') do set "BASEPY=%%P"
)
if not defined BASEPY if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No hay Python instalado. Instala Anaconda o Python 3.11+ y volve a ejecutar este archivo.
    set "_COD=1"
    goto :salir
)
if defined BASEPY echo Python base: %BASEPY%

:: .venv con acceso a los paquetes ya instalados (Anaconda trae pandas, pywin32, etc.)
if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno .venv ...
    "%BASEPY%" -m venv --system-site-packages .venv
    if errorlevel 1 echo [AVISO] No se pudo crear .venv. Se instala sobre el Python base.
)
call "%~dp0scripts\_python.bat"
if errorlevel 1 (
    set "_COD=1"
    goto :salir
)
echo Python a usar: %PY%
echo.

"%PY%" -m pip install --disable-pip-version-check --no-warn-script-location -q -r "Proyecto_facturas_v2\requirements.txt"
if errorlevel 1 (
    echo.
    echo [ERROR] Fallo la instalacion de dependencias. Revisa la conexion a internet / proxy y reintenta.
    set "_COD=1"
    goto :salir
)
echo Dependencias principales: OK
echo Instalando OCR opcional ^(si falla, el sistema funciona igual sin OCR^)...
"%PY%" -m pip install --disable-pip-version-check --no-warn-script-location -q -r "Proyecto_facturas_v2\requirements-opcional.txt" >nul 2>&1 && echo OCR: OK || echo OCR: no instalado ^(opcional^)

echo.
echo Verificando...
"%PY%" -W ignore -m unittest discover -s "%~dp0Proyecto_facturas_v2\tests" -q
if errorlevel 1 (
    echo [ERROR] Algun test fallo. Ver mensajes arriba.
    set "_COD=1"
    goto :salir
)
echo.
echo ======================================================
echo    INSTALACION OK. Usa MENU.bat para operar.
echo ======================================================

:salir
if defined _PAUSA pause
exit /b %_COD%
