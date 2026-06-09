@echo off
title Auto Product Listing Launcher
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

echo ============================================================
echo           Auto Product Listing Launcher (Auto Listing)
echo ============================================================
echo.

:: Check local virtual environment
if exist ".venv\Scripts\activate.bat" (
    echo [INFO] Activating virtual environment...
    call ".venv\Scripts\activate.bat"
) else (
    echo [INFO] No local virtual environment found, using system Python...
)

echo [INFO] Starting main script auto_listing.py...
echo ------------------------------------------------------------
python auto_listing.py
echo ------------------------------------------------------------

echo.
if %errorlevel% neq 0 (
    echo [ERROR] Process exited with error.
) else (
    echo [INFO] Process completed successfully.
)
echo.
pause
