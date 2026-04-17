@echo off
chcp 65001 >nul
title ZTE-AIS Gulf Solar BTS Dashboard

echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║   ZTE-AIS Gulf Solar BTS 2025 Dashboard         ║
echo  ║   Starting local server on http://127.0.0.1:8000 ║
echo  ╚══════════════════════════════════════════════════╝
echo.

:: Change to the directory where this .bat file is located
cd /d "%~dp0"

:: Check if .venv exists
if not exist ".venv\Scripts\python.exe" (
    echo  [ERROR] Virtual environment not found!
    echo  Please run setup.bat first or check the .venv folder.
    echo.
    pause
    exit /b 1
)

:: Start FastAPI server in background
echo  [1/2] Starting FastAPI server...
start "" ".venv\Scripts\python.exe" -m uvicorn main:app --host 127.0.0.1 --port 8000 --reload

:: Wait for server to be ready
echo  [2/2] Waiting for server to start...
timeout /t 3 /nobreak >nul

:: Open browser automatically
echo  Opening browser...
start "" "http://127.0.0.1:8000"

echo.
echo  ✓ Dashboard is running at http://127.0.0.1:8000
echo  ✓ Browser should open automatically.
echo.
echo  ┌─────────────────────────────────────────────────┐
echo  │  To STOP the server:                            │
echo  │  Close this window OR press any key below       │
echo  │  (this will shut down the server process)       │
echo  └─────────────────────────────────────────────────┘
echo.
pause

:: Kill uvicorn when user presses a key
echo  Stopping server...
taskkill /f /im python.exe >nul 2>&1
echo  Server stopped. Goodbye!
timeout /t 2 /nobreak >nul
