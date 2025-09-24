@echo off
chcp 65001 >nul
echo ==========================================
echo      Running Yandex Wordstat parser
echo ==========================================
echo.

if not exist ".venv" (
    echo ERROR: Virtual environment not found!
    echo Run install.bat for installation
    pause
    exit /b 1
)

if not exist "queries.txt" (
    echo ERROR: File queries.txt not found!
    echo Create queries.txt file and add search queries
    pause
    exit /b 1
)

echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo.
echo Starting parser...
python wordstat_parser.py

echo.
echo Done! Check the result in wordstat_report.xlsx file
pause
