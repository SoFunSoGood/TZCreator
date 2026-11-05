# QWEN.md — инструкция для Qwen Code

## О проекте
Генератор заявок кандидатов (docx) из входных xlsx-файлов. Данные кандидатов читаются
из `input/*.xlsx`, подставляются в шаблон `templates/заявка_шаблон.docx` (docxtpl/Jinja2),
результат складывается в `output/`. Константы (номера, даты, ставки) — в `config/config.yaml`.

## Структура
- `config/config.yaml` — переменные: номера, даты, ставки. НЕ хранить в коде.
- `templates/` — эталонный docx с Jinja2-плейсхолдерами вида `{{ field }}`, `{% for %}`.
- `reference/` — примеры готовых заявок 106/110 только для сверки; не менять.
- `input/` — сюда кладём 1–20 xlsx кандидатов (не в VCS).
- `output/` — сюда кладём результат (не в VCS).
- `src/models.py` — dataclass-структуры данных кандидата и заявки.
- `src/parser.py` — чтение/валидация xlsx через openpyxl.
- `src/generator.py` — агрегация данных + рендер docx через docxtpl.
- `src/main.py` — CLI-точка входа (argparse).
- `tests/` — pytest; запуск: `pytest -v`.

## Команды
- Установка окружения: `bash setup.sh` (Linux/macOS) или `setup.ps1` (Windows).
- Тесты: `pytest -v`
- Запуск: `python -m src.main --input input --config config/config.yaml --output output`

## Правила для Qwen Code
1. Python 3.12, зависимости только из `requirements.txt`; новые пакеты согласовывать.
2. Не трогать файлы в `reference/` и сам эталонный шаблон без явного запроса.
3. Персональные данные кандидатов не логировать и не коммитить (`input/`, `output/` в .gitignore).
4. Любое изменение логики парсинга/генерации сопровождать тестом в `tests/`.
5. Перед коммитом прогонять `pytest -v` — CI запускает тесты автоматически.
6. Стилистика: snake_case, type hints, докстринги на русском допустимы.
