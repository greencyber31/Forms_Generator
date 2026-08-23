@echo off
title PCIC Form Studio Launcher
cd /d "%~dp0"

echo ===================================================
echo           PCIC Form Studio - Web Suite
echo ===================================================
echo.

:: Detect Python executable
set "PYTHON_EXE="

:: 1. Check local virtual environment if present
if exist "venv\Scripts\python.exe" (
    set "PYTHON_EXE=venv\Scripts\python.exe"
) else if exist ".venv\Scripts\python.exe" (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
)

:: 2. Check system Python / py launcher
if "%PYTHON_EXE%"=="" (
    where python.exe >nul 2>nul
    if %errorlevel% equ 0 (
        python.exe -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>nul
        if %errorlevel% equ 0 set "PYTHON_EXE=python.exe"
    )
)

:: 3. Check AppData Python312
if "%PYTHON_EXE%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    )
)

:: 4. Check AppData Python311
if "%PYTHON_EXE%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    )
)

:: If still not found, notify user
if "%PYTHON_EXE%"=="" (
    echo [ERROR] Python 3 was not detected on your system.
    echo Please install Python 3.10+ from https://www.python.org or Microsoft Store.
    pause
    exit /b 1
)

echo [INFO] Using Python: %PYTHON_EXE%

:: Ensure requirements are installed
echo [INFO] Verifying dependencies...
"%PYTHON_EXE%" -c "import flask, pandas, openpyxl, docxtpl, docx, pypdf, win32com.client" >nul 2>nul
if %errorlevel% neq 0 (
    echo [INFO] Installing required dependencies...
    "%PYTHON_EXE%" -m pip install -r requirements.txt
)

echo.
echo [INFO] Starting PCIC Form Studio at http://localhost:5000
echo [INFO] Opening web browser...
start "" "http://localhost:5000"

:: Start the Flask app
"%PYTHON_EXE%" app.py
pause
