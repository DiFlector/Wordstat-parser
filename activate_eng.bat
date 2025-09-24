@echo off
chcp 65001 >nul
echo Activating virtual environment .venv...

if not exist ".venv" (
    echo ERROR: Virtual environment not found!
    echo Run install.bat to create virtual environment
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
echo ✓ Virtual environment activated

echo.
echo Available commands:
echo   python wordstat_parser.py     - run parser
echo   python example_usage.py       - demonstration
echo   pip list                      - list installed packages
echo   deactivate                    - exit virtual environment
echo.
