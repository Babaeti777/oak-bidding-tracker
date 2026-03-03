@echo off
title OAK BUILDERS - Bidding Tracker Web App
color 0A
cd /d "%~dp0"

echo.
echo   ============================================================
echo     OAK BUILDERS - Bidding Tracker (Web App)
echo   ============================================================
echo.

REM ── Check Python ──
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo    Python is not installed or not in PATH.
    echo    Download from https://python.org
    pause
    exit /b 1
)

REM ── Install dependencies if needed ──
python -c "import flask" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo    Installing required packages...
    pip install flask flask-login openpyxl -q
)

REM ── Get local IP for phone access ──
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /C:"IPv4 Address"') do (
    set IP=%%a
    goto :found_ip
)
:found_ip
set IP=%IP: =%

echo    Starting web server...
echo.
echo    ────────────────────────────────────────
echo    Open in browser:
echo.
echo      PC:    http://localhost:5000
echo      Phone: http://%IP%:5000
echo.
echo    ────────────────────────────────────────
echo    Press Ctrl+C to stop the server.
echo.

python app.py
pause
