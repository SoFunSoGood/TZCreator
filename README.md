# Генератор заявок кандидатов

Проект автоматизирует формирование заявок кандидатов (`.docx`) из табличных
файлов `.xlsx`: данные читаются, валидируются, агрегируются и подставляются
в эталонный шаблон Word с Jinja2-плейсхолдерами.

## Структура репозитория

```
.
├── .github/workflows/ci.yml      # CI: автозапуск тестов
├── .gitignore
├── QWEN.md                        # инструкция для Qwen Code
├── README.md
├── requirements.txt
├── setup.sh / setup.ps1           # скрипты установки окружения
├── config/
│   └── config.yaml                # переменные: номера, даты, ставки
├── templates/
│   └── заявка_шаблон.docx         # эталон с Jinja2-плейсхолдерами
├── reference/                     # примеры 106/110 (для сверки)
├── input/                         # сюда кладём 1–20 xlsx кандидатов
├── output/                        # сюда кладём результат
├── src/
│   ├── __init__.py
│   ├── models.py                  # структуры данных
│   ├── parser.py                  # чтение xlsx
│   ├── generator.py               # агрегация + рендер docx
│   └── main.py                    # CLI-точка входа
└── tests/
    └── test_parser.py
```

## Быстрый старт

### Linux / macOS
```bash
bash setup.sh
source .venv/bin/activate
```

### Windows (PowerShell)
```powershell
.\setup.ps1
.\.venv\Scripts\Activate.ps1
```

## Использование

1. Положите 1–20 файлов `.xlsx` кандидатов в папку `input/`.
2. При необходимости отредактируйте `config/config.yaml` (номера, даты, ставки).
3. Запустите:

```bash
python -m src.main --input input --config config/config.yaml --output output
```

Готовые заявки появятся в `output/`.

## Тесты

```bash
pytest -v
```

CI (GitHub Actions) автоматически прогоняет тесты при каждом push/PR в `main`.
