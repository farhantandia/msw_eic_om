@echo off
title Outage EIC Work Order Monitoring System
echo ==================================================================
echo   Starting Outage EIC Work Order Monitoring System...
echo ==================================================================
echo.
echo Opening Web Browser at http://localhost:8000 ...
start http://localhost:8000
echo.
set PYTHON_CMD=python
if exist "D:\miniconda3\python.exe" (
    set PYTHON_CMD="D:\miniconda3\python.exe"
)

%PYTHON_CMD% server.py
if errorlevel 1 (
    echo.
    echo [Fallback] Menjalankan server.exe...
    server.exe
)
pause

