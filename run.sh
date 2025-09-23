#!/bin/bash

echo "=========================================="
echo "     Запуск парсера Яндекс Вордстат"
echo "=========================================="
echo

if [ ! -d ".venv" ]; then
    echo "ОШИБКА: Виртуальное окружение не найдено!"
    echo "Запустите ./install.sh для установки"
    exit 1
fi

if [ ! -f "queries.txt" ]; then
    echo "ОШИБКА: Файл queries.txt не найден!"
    echo "Создайте файл queries.txt и добавьте поисковые запросы"
    exit 1
fi

echo "Активируем виртуальное окружение..."
source .venv/bin/activate

echo
echo "Запускаем парсер..."
python wordstat_parser.py

echo
echo "Готово! Проверьте результат в файле wordstat_report.xlsx"
