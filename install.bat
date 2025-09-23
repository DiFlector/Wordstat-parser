@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
echo ==========================================
echo    Установка зависимостей для парсера
echo         Яндекс Вордстат
echo ==========================================
echo.

echo Проверяем наличие Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ОШИБКА: Python не найден!
    echo Установите Python с официального сайта: https://www.python.org/
    pause
    exit /b 1
)

echo ✓ Python найден
echo.

echo Удаляем старое виртуальное окружение если есть проблемы...
if exist ".venv" (
    echo Обнаружено существующее виртуальное окружение
    choice /M "Пересоздать виртуальное окружение"
    if !errorlevel!==1 (
        rmdir /s /q .venv
        echo ✓ Старое окружение удалено
    )
)

echo Создаем виртуальное окружение .venv...
if not exist ".venv" (
    python -m venv .venv
    if %errorlevel% neq 0 (
        echo ОШИБКА: Не удалось создать виртуальное окружение!
        pause
        exit /b 1
    )
    echo ✓ Виртуальное окружение создано
) else (
    echo ✓ Виртуальное окружение уже существует
)
echo.

echo Активируем виртуальное окружение...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось активировать виртуальное окружение!
    pause
    exit /b 1
)

echo Устанавливаем зависимости в виртуальное окружение...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ОШИБКА: Не удалось установить зависимости!
    echo Попробуйте установить вручную: pip install requests selenium beautifulsoup4 openpyxl webdriver-manager
    pause
    exit /b 1
)

echo.
echo ✓ Установка завершена!
echo.
echo Для запуска программы используйте:
echo   activate.bat          (активация виртуального окружения)
echo   python wordstat_parser.py
echo.
echo Или просто:
echo   run.bat               (запуск с автоматической активацией)
echo.
pause
