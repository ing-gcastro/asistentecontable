@echo off
:: _python.bat — Busca el Python a usar y lo deja en la variable PY.
:: Uso desde otro .bat:   call "%~dp0scripts\_python.bat" || exit /b 1
:: Orden: .venv del repo > Anaconda del usuario > py launcher > python del PATH.
set "PY="
set "_RAIZ=%~dp0.."

if exist "%_RAIZ%\.venv\Scripts\python.exe" set "PY=%_RAIZ%\.venv\Scripts\python.exe"
if not defined PY if exist "%USERPROFILE%\AppData\Local\anaconda3\python.exe" set "PY=%USERPROFILE%\AppData\Local\anaconda3\python.exe"
if not defined PY if exist "%USERPROFILE%\anaconda3\python.exe" set "PY=%USERPROFILE%\anaconda3\python.exe"
if not defined PY if exist "%ProgramData%\anaconda3\python.exe" set "PY=%ProgramData%\anaconda3\python.exe"
if not defined PY if exist "%USERPROFILE%\miniconda3\python.exe" set "PY=%USERPROFILE%\miniconda3\python.exe"
if not defined PY (
    py -3 -c "import sys" >nul 2>&1 && for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)"') do set "PY=%%P"
)
if not defined PY (
    python -c "import sys" >nul 2>&1 && for /f "delims=" %%P in ('python -c "import sys;print(sys.executable)"') do set "PY=%%P"
)

if not defined PY (
    echo.
    echo [ERROR] No se encontro Python en esta PC.
    echo         Instala Anaconda o Python 3.11+ desde https://www.python.org/downloads/
    echo         ^(marcar "Add python.exe to PATH"^) y despues ejecuta INSTALAR.bat
    echo.
    exit /b 1
)
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
set "PYTHONWARNINGS=ignore::FutureWarning,ignore::DeprecationWarning"
exit /b 0
