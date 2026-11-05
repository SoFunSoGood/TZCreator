"""Чтение xlsx-файлов кандидатов."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Iterable, Optional

from openpyxl import load_workbook

from src.models import Candidate

# Ожидаемые заголовки колонок (регистронезависимо, допускаются синонимы)
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "last_name": ("фамилия", "фэмилия", "last name", "surname"),
    "first_name": ("имя", "first name"),
    "middle_name": ("отчество", "middle name", "patronymic"),
    "full_name": ("фио", "full name"),
    "birth_date": ("дата рождения", "др", "birth date"),
    "position": ("должность", "position", "rank"),
    "role": ("роль", "role"),
    "unit": ("подразделение", "часть", "unit", "department"),
    "service_name": ("наименование услуг", "наименование услуги", "услуга"),
    "hours": ("объем работ", "объём работ", "час", "hours"),
    "hour_rate": ("стоимость часа",),
    "cost": ("стоимость услуги",),
}


def _norm(value: object) -> str:
    return str(value).strip().lower() if value is not None else ""


def _clean(value: object) -> str:
    """Значение ячейки без изменения регистра (для ФИО и должностей)."""
    return str(value).strip() if value is not None else ""


def _map_headers(header_row: Iterable[object]) -> dict[str, int]:
    """Сопоставление имён полей с индексами колонок по строке заголовков."""
    mapping: dict[str, int] = {}
    for idx, cell in enumerate(header_row):
        name = _norm(cell)
        for field_name, aliases in COLUMN_ALIASES.items():
            if any(name == a or name.startswith(a) for a in aliases):
                mapping.setdefault(field_name, idx)
    return mapping


def _to_date(value: object) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        try:
            return datetime.strptime(value.strip(), "%d.%m.%Y").date()
        except ValueError:
            return None
    return None


def _to_float(value: object) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and value.strip():
        try:
            return float(value.strip().replace(",", ".").replace(" ", ""))
        except ValueError:
            return 0.0
    return 0.0


def _split_full_name(text: str) -> tuple[str, str, Optional[str]]:
    """Разбить «Фамилия И.О.» / «Фамилия Имя Отчество» на части."""
    parts = text.split()
    if not parts:
        return "", "", None
    last = parts[0]
    first = parts[1] if len(parts) > 1 else ""
    middle = " ".join(parts[2:]) if len(parts) > 2 else None
    return last, first, middle or None


def parse_candidate_file(path: Path, header_row: int = 1,
                         sheet_name: Optional[str] = None) -> list[Candidate]:
    """Прочитать всех кандидатов из одного xlsx-файла.

    Поддерживаются два формата:
    1) Классический: колонки Фамилия/Имя (+Отчество, Дата рождения и т.д.).
    2) Шаблон заявки кандидата: ФИО / Роль / Объем работ / Стоимость часа /
       Стоимость услуги — несколько строк услуг на одного кандидата.
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    mapping: dict[str, int] = {}
    rows_iter = enumerate(ws.iter_rows(values_only=True), start=1)
    by_name: dict[str, Candidate] = {}

    for rownum, row in rows_iter:
        if row is None or all(c in (None, "") for c in row):
            continue

        # Пока не нашли строку заголовков — ищем её
        if not mapping:
            candidate_map = _map_headers(row)
            if {"last_name", "first_name"} <= candidate_map.keys() or \
               {"full_name", "role"} <= candidate_map.keys():
                mapping = candidate_map
            continue

        def get(key: str) -> object:
            idx = mapping.get(key)
            return row[idx] if idx is not None and idx < len(row) else None

        last = _clean(get("last_name"))
        first = _clean(get("first_name"))
        if not (last and first):
            fio_text = _clean(get("full_name"))
            if not fio_text or fio_text.lower().startswith("итог"):
                continue
            last, first, mid = _split_full_name(fio_text)
            if not (last and first):
                continue
            if mid:
                middle = mid
            else:
                middle = _clean(get("middle_name")) or None
        else:
            middle = _clean(get("middle_name")) or None

        key = f"{last}|{first}|{middle or ''}"
        candidate = by_name.get(key)
        if candidate is None:
            candidate = Candidate(
                last_name=last,
                first_name=first,
                middle_name=middle,
                birth_date=_to_date(get("birth_date")),
                position=_clean(get("position")),
                unit=_clean(get("unit")),
                role=_clean(get("role")),
                source_file=path.name,
            )
            by_name[key] = candidate

        service_name = _clean(get("service_name"))
        if service_name:
            from src.models import Service
            candidate.services.append(Service(
                name=service_name,
                hours=_to_float(get("hours")),
                hour_rate=_to_float(get("hour_rate")),
                cost=_to_float(get("cost")),
            ))

    wb.close()
    if not mapping:
        raise ValueError(f"В файле {path.name} не найдена строка заголовков")
    if not by_name:
        raise ValueError(
            f"В файле {path.name} не найдены колонки Фамилия/Имя или ФИО/Роль"
        )
    return list(by_name.values())


def parse_input_dir(input_dir: Path, header_row: int = 1,
                    max_files: int = 20) -> list[Candidate]:
    """Прочитать все xlsx-кандидаты из каталога input/ (до max_files)."""
    files = sorted(p for p in input_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    if len(files) > max_files:
        raise ValueError(f"Файлов больше {max_files}: {len(files)}")
    result: list[Candidate] = []
    for path in files:
        result.extend(parse_candidate_file(path, header_row=header_row))
    return result
