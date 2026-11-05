"""Тесты парсера xlsx."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from src.parser import _map_headers, _to_date, parse_candidate_file, parse_input_dir


def make_xlsx(path: Path, rows: list[list]) -> Path:
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


HEADERS = ["Фамилия", "Имя", "Отчество", "Дата рождения", "Должность", "Подразделение"]


def test_map_headers():
    mapping = _map_headers(HEADERS)
    assert mapping["last_name"] == 0
    assert mapping["first_name"] == 1
    assert mapping["middle_name"] == 2
    assert mapping["birth_date"] == 3
    assert mapping["position"] == 4
    assert mapping["unit"] == 5


def test_to_date_variants():
    assert _to_date(datetime(2000, 1, 2)) == date(2000, 1, 2)
    assert _to_date(date(1999, 5, 6)) == date(1999, 5, 6)
    assert _to_date("07.08.1990") == date(1990, 8, 7)
    assert _to_date("") is None
    assert _to_date(None) is None
    assert _to_date("не дата") is None


def test_parse_candidate_file(tmp_path: Path):
    f = make_xlsx(tmp_path / "cand.xlsx", [
        HEADERS,
        ["Иванов", "Иван", "Иванович", datetime(1995, 3, 10), "младший сержант", "П4"],
        ["Петров", "Пётр", None, "12.04.1996", "рядовой", "П5"],
        [None, None, None, None, None, None],  # пустая строка игнорируется
    ])
    candidates = parse_candidate_file(f)
    assert len(candidates) == 2
    assert candidates[0].full_name == "Иванов Иван Иванович"
    assert candidates[0].birth_date == date(1995, 3, 10)
    assert candidates[1].last_name == "Петров"  # регистр сохраняется
    assert candidates[1].middle_name is None


def test_parse_missing_columns_raises(tmp_path: Path):
    f = make_xlsx(tmp_path / "bad.xlsx", [["A", "B"], ["1", "2"]])
    with pytest.raises(ValueError):
        parse_candidate_file(f)


def test_parse_input_dir(tmp_path: Path):
    make_xlsx(tmp_path / "01_a.xlsx", [HEADERS, ["А", "Б", None, None, None, None]])
    make_xlsx(tmp_path / "02_b.xlsx", [HEADERS, ["В", "Г", None, None, None, None]])
    candidates = parse_input_dir(tmp_path)
    assert len(candidates) == 2


def test_parse_input_dir_too_many_files(tmp_path: Path):
    for i in range(3):
        make_xlsx(tmp_path / f"c{i}.xlsx", [HEADERS, ["X", "Y", None, None, None, None]])
    with pytest.raises(ValueError):
        parse_input_dir(tmp_path, max_files=2)


CANDIDATE_HEADERS = ["", "Наименование Услуг", "ФИО", "Роль",
                     "Объем работ, ч/час", "Стоимость часа, руб.*",
                     "Стоимость услуги, руб."]


def test_parse_candidate_template_format(tmp_path: Path):
    """Формат «Шаблон заявки кандидата»: ФИО + несколько строк услуг."""
    f = make_xlsx(tmp_path / "tpl.xlsx", [
        ["", "Кандидат на выполнение работ по проекту"],
        CANDIDATE_HEADERS,
        ["", "Услуга 1", "Алаткин А.", "системный аналитик", 38, 1529.26, 58111.88],
        ["", "Услуга 2", "Алаткин А.", "системный аналитик", 38, 1529.26, 58111.88],
        ["", "Итого:", "", "", "", "", 116223.76],
    ])
    candidates = parse_candidate_file(f)
    assert len(candidates) == 1
    c = candidates[0]
    assert c.full_name == "Алаткин А."
    assert c.role == "системный аналитик"
    assert len(c.services) == 2
    assert c.total_cost == 116223.76
