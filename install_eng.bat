@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
echo ==========================================
echo    Installing dependencies for Yandex
echo           Wordstat parser
echo ==========================================
echo.

echo Checking for Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    echo Install Python from the official website: https://www.python.org/
    pause
    exit /b 1
)

echo ✓ Python found
echo.

echo Removing old virtual environment if there are issues...
if exist ".venv" (
    echo Found existing virtual environment
    choice /M "Recreate virtual environment"
    if !errorlevel!==1 (
        rmdir /s /q .venv
        echo ✓ Old environment removed
    )
)

echo Creating virtual environment .venv...
if not exist ".venv" (
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment!
        pause
        exit /b 1
    )
    echo ✓ Virtual environment created
) else (
    echo ✓ Virtual environment already exists
)
echo.

echo Activating virtual environment...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)

echo Installing dependencies in virtual environment...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies!
    echo Try installing manually: pip install requests selenium beautifulsoup4 openpyxl webdriver-manager
    pause
    exit /b 1
)

echo.
echo ✓ Installation completed!
echo.
echo To run the program use:
echo   activate.bat          (activate virtual environment)
echo   python wordstat_parser.py
echo.
echo Or simply:
echo   run.bat               (run with automatic activation)
echo.
pause
