@echo off
chcp 65001 >nul
title Auto Downloader
call "%~dp0..\scripts\_python.bat" || (pause & exit /b 1)
cd /d "%~dp0"
"%PY%" auto_downloader.py %*
