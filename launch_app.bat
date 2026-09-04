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

:: 2. Check py launcher (Python Launcher for Windows)
if "%PYTHON_EXE%"=="" (
    where py.exe >nul 2>nul
    if not errorlevel 1 (
        for /f "delims=" %%I in ('py -3 -c "import sys, os; print(os.path.abspath(sys.executable))" 2^>nul') do (
            if exist "%%I" set "PYTHON_EXE=%%I"
        )
    )
)

:: 3. Check AppData standard Python installations
if "%PYTHON_EXE%"=="" (
    if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    ) else if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    )
)

:: 4. Check ProgramFiles / system installations
if "%PYTHON_EXE%"=="" (
    if exist "%ProgramFiles%\Python313\python.exe" (
        set "PYTHON_EXE=%ProgramFiles%\Python313\python.exe"
    ) else if exist "%ProgramFiles%\Python312\python.exe" (
        set "PYTHON_EXE=%ProgramFiles%\Python312\python.exe"
    ) else if exist "%ProgramFiles%\Python311\python.exe" (
        set "PYTHON_EXE=%ProgramFiles%\Python311\python.exe"
    ) else if exist "%ProgramFiles%\Python310\python.exe" (
        set "PYTHON_EXE=%ProgramFiles%\Python310\python.exe"
    )
)

:: 5. Test system PATH python executable (verifying it is real and not Microsoft Store stub)
if "%PYTHON_EXE%"=="" (
    for /f "delims=" %%I in ('python -c "import sys, os; print(os.path.abspath(sys.executable))" 2^>nul') do (
        if exist "%%I" set "PYTHON_EXE=%%I"
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
"%PYTHON_EXE%" -c "import flask, pandas, openpyxl, docxtpl, docx, pypdf, win32com.client, docxcompose" >nul 2>nul
if errorlevel 1 (
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
