#!/bin/bash

echo "=========================================="
echo "    Установка зависимостей для парсера"
echo "         Яндекс Вордстат"
echo "=========================================="
echo

# Проверяем наличие Python
if ! command -v python3 &> /dev/null; then
    echo "ОШИБКА: Python3 не найден!"
    echo "Установите Python3:"
    echo "Ubuntu/Debian: sudo apt install python3 python3-pip python3-venv"
    echo "CentOS/RHEL: sudo yum install python3 python3-pip"
    echo "macOS: brew install python3"
    exit 1
fi

echo "✓ Python3 найден"
echo

# Создаем виртуальное окружение
if [ ! -d ".venv" ]; then
    echo "Создаем виртуальное окружение .venv..."
    python3 -m venv .venv
    echo "✓ Виртуальное окружение создано"
else
    echo "✓ Виртуальное окружение уже существует"
fi
echo

# Активируем виртуальное окружение
echo "Активируем виртуальное окружение..."
source .venv/bin/activate

echo "Обновляем pip в виртуальном окружении..."
python -m pip install --upgrade pip

echo
echo "Устанавливаем зависимости в виртуальное окружение..."
pip install -r requirements.txt

echo
echo "✓ Установка завершена!"
echo
echo "Для запуска программы используйте:"
echo "  source .venv/bin/activate    (активация виртуального окружения)"
echo "  python wordstat_parser.py"
echo
echo "Или просто:"
echo "  ./run.sh                     (запуск с автоматической активацией)"
echo
