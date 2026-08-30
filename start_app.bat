@echo off
title Outage EIC Work Order Monitoring System
echo ==================================================================
echo   Starting Outage EIC Work Order Monitoring System...
echo ==================================================================
echo.
echo Opening Web Browser at http://localhost:8000 ...
start http://localhost:8000
echo.
python server.py
if errorlevel 1 (
    echo.
    echo [Fallback] Menjalankan server.exe...
    server.exe
)
pause
