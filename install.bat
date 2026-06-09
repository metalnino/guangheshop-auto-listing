@echo off
title Auto Product Listing Installer
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

echo ============================================================
echo           Auto Product Listing - Installer
echo ============================================================
echo.
echo   This script will:
echo     1. Create a local virtual environment (.venv)
echo     2. Upgrade pip and install dependencies (requirements.txt)
echo     3. Create required directories (pending_upload, completed_upload, logs)
echo.
echo   Press any key to start...
pause >nul

echo.
echo [INFO] Searching for Python...

:: Find Python executable (skip Windows Store stub)
set "PY_EXE="
for /f "delims=" %%P in ('where python 2^>nul') do (
    echo %%P | findstr /i "WindowsApps" >nul
    if errorlevel 1 (
        set "PY_EXE=%%P"
        goto :py_found
    )
)
:py_found

if "%PY_EXE%"=="" (
    echo [ERROR] Python was not found in your PATH.
    echo         Please install Python 3.10+ manually and check "Add to PATH".
    goto :fail
)

echo [INFO] Using Python: %PY_EXE%

:: Create .venv if not exists
if not exist ".venv" (
    echo [INFO] Creating virtual environment (.venv)...
    "%PY_EXE%" -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        goto :fail
    )
) else (
    echo [INFO] Virtual environment (.venv) already exists.
)

:: Get path to venv python
set "VENV_PY=%CD%\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment python executable not found.
    goto :fail
)

echo.
echo [INFO] Upgrading pip...
"%VENV_PY%" -m pip install -U pip >nul 2>&1

echo [INFO] Installing packages from requirements.txt...
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Package installation failed. Please check your internet connection.
    goto :fail
)

:: Create directories
echo.
echo [INFO] Creating required folders...
if not exist "pending_upload" (
    mkdir "pending_upload"
    echo   Created folder: pending_upload
)
if not exist "completed_upload" (
    mkdir "completed_upload"
    echo   Created folder: completed_upload
)
if not exist "logs" (
    mkdir "logs"
    echo   Created folder: logs
)

echo.
echo ============================================================
echo   [OK] Installation completed successfully!
echo   You can now run "run_auto_listing.bat" to start.
echo ============================================================
goto :end

:fail
echo.
echo [ERROR] Installation failed.
echo.

:end
pause
