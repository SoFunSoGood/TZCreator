#!/usr/bin/env bash
# setup.sh — установка окружения проекта (Linux/macOS)
set -euo pipefail

cd "$(dirname "$0")"

PYTHON=${PYTHON:-python3}

echo "==> Создаю виртуальное окружение .venv"
"$PYTHON" -m venv .venv

# shellcheck disable=SC1091
source .venv/bin/activate

echo "==> Обновляю pip и ставлю зависимости"
python -m pip install --upgrade pip
pip install -r requirements.txt

echo "==> Проверяю тесты"
pytest -v || echo "ВНИМАНИЕ: тесты завершились с ошибками"

echo "Готово. Активация окружения: source .venv/bin/activate"
