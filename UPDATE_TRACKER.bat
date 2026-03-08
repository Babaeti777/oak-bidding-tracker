@echo off
title OAK BUILDERS - Bidding Tracker Update
color 0F
cd /d "%~dp0"

echo.
echo   ============================================================
echo     OAK BUILDERS - Bidding Tracker Update
echo   ============================================================
echo.

REM ── Check if Python is available ──
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo    Python is not installed or not in PATH.
    echo    Download from https://python.org
    echo    Make sure "Add Python to PATH" is checked during install.
    echo.
    pause
    exit /b 1
)

REM ── Check if tracker Excel is open (lock file exists) ──
if exist ".~lock.Bidding_Tracker_Pro_v4.xlsx#" (
    echo    WARNING: The tracker appears to be open in Excel.
    echo    Close it first for best results, or press any key to try anyway...
    echo.
    pause >nul
)

REM ── Run the update script ──
python update_tracker.py
set EXITCODE=%ERRORLEVEL%

if %EXITCODE% NEQ 0 (
    echo.
    echo   ────────────────────────────────────────────────────────
    echo    The script exited with an error (code %EXITCODE%).
    echo    If the issue persists, share the error above with Claude.
    echo   ────────────────────────────────────────────────────────
    echo.
)

echo.
echo   Press any key to close...
pause >nul
