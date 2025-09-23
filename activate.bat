@echo off
chcp 65001 >nul
echo Активируем виртуальное окружение .venv...

if not exist ".venv" (
    echo ОШИБКА: Виртуальное окружение не найдено!
    echo Запустите install.bat для создания виртуального окружения
    pause
    exit /b 1
)

call .venv\Scripts\activate.bat
echo ✓ Виртуальное окружение активировано

echo.
echo Доступные команды:
echo   python wordstat_parser.py     - запуск парсера
echo   python example_usage.py       - демонстрация
echo   pip list                      - список установленных пакетов
echo   deactivate                    - выход из виртуального окружения
echo.
