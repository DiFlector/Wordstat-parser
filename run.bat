@echo off
chcp 65001 >nul
echo ==========================================
echo      Запуск парсера Яндекс Вордстат
echo ==========================================
echo.

if not exist ".venv" (
    echo ОШИБКА: Виртуальное окружение не найдено!
    echo Запустите install.bat для установки
    pause
    exit /b 1
)

if not exist "queries.txt" (
    echo ОШИБКА: Файл queries.txt не найден!
    echo Создайте файл queries.txt и добавьте поисковые запросы
    pause
    exit /b 1
)

echo Активируем виртуальное окружение...
call .venv\Scripts\activate.bat

echo.
echo Запускаем парсер...
python wordstat_parser.py

echo.
echo Готово! Проверьте результат в файле wordstat_report.xlsx
pause
