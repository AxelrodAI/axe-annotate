@echo off
TITLE Axe Annotate Launcher
echo ========================================================
echo               Axe Annotate v2.2 Launcher
echo                   (Clean Edition)
echo ========================================================
echo.

set PYTHON_CMD=python

:: 1. Check for Python in PATH
%PYTHON_CMD% --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [Check] 'python' command not found. Checking standard paths...
    
    :: Check typical User install location for Python 3.12
    if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe" (
        set PYTHON_CMD="%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe"
        echo [Found] Python 3.12 found.
    ) else (
        echo [ERROR] Python not found.
        echo Install from https://www.python.org/downloads/
        echo IMPORTANT: Check "Add Python to PATH" during install.
        pause
        exit /b
    )
)

echo Using Python: %PYTHON_CMD%

:: 2. Install Requirements
echo.
echo [2/3] Installing dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b
)
echo [2/3] Dependencies OK.

:: 3. Quick Connection Test
echo.
echo [3/3] Quick Excel connection test...
%PYTHON_CMD% tests\stress_test_excel.py --quick
echo.

:: 4. Run the Tool
echo ========================================================
echo                     STARTING TOOL
echo ========================================================
echo.
echo SHORTCUTS:
echo   Ctrl+Shift+m   Auto-Annotate selected cell
echo   Ctrl+Shift+2   Custom Prompt + Annotate
echo   Ctrl+Shift+h   Check Excel connection
echo   Esc            Quit
echo.
echo TIP: Press Esc in Excel if editing a cell.
echo.
%PYTHON_CMD% main.py
pause
:: end of file
