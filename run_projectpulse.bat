@echo off
title ProjectPulse

cd /d "%~dp0"

echo Starting ProjectPulse...

start "" "C:\Program Files\Anaconda3\python.exe" server.py

timeout /t 3 /nobreak >nul

start "" http://127.0.0.1:8000

exit