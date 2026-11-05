"""Корневой conftest: добавляет корень проекта в sys.path,
чтобы тесты могли импортировать пакет `src` при запуске pytest из CI."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))