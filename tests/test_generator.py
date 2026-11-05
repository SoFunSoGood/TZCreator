"""Тесты генератора: агрегация кандидатов и рендер заявки docx по шаблону."""
from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
from docxtpl import DocxTemplate
from openpyxl import Workbook

from src.generator import build_application, generate, render_application
from src.models import Candidate, Service

TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "заявка_шаблон.docx"


@pytest.fixture(scope="module")
def candidates() -> list[Candidate]:
    return [
        Candidate(
            last_name="Иванов", first_name="Иван", middle_name="Иванович",
            birth_date=date(1990, 1, 15), role="аналитик",
            services=[Service(name="Анализ", hours=10, hour_rate=1000.0, cost=10000.0)],
        ),
        Candidate(
            last_name="Петров", first_name="Пётр",
            services=[
                Service(name="Разработка", hours=20, hour_rate=2000.0, cost=40000.0),
                Service(name="Тестирование", hours=5, hour_rate=1500.0, cost=7500.0),
            ],
        ),
    ]


def test_build_application(candidates: list[Candidate]) -> None:
    app = build_application(candidates, number=106, created=date(2026, 11, 5))
    assert app.number == 106
    assert app.title == "Заявка 106"
    assert len(app.candidates) == 2


def test_render_application(candidates: list[Candidate], tmp_path: Path) -> None:
    """Рендер реального эталонного шаблона: плейсхолдеры подставлены, таблица развернута."""
    if not TEMPLATE.is_file():
        pytest.skip("Эталонный шаблон отсутствует")
    app = build_application(candidates, number=106, created=date(2026, 11, 5))
    out = render_application(app, TEMPLATE, tmp_path)
    assert out.is_file() and out.suffix == ".docx"

    # проверка результата: читаем document.xml и убираем XML-теги
    import re
    import zipfile
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    plain = re.sub(r"<[^>]+>", "", xml)
    assert "Иванов Иван Иванович" in plain
    assert "Тестирование" in plain
    assert "{{" not in plain  # все плейсхолдеры отрендерены
    assert "57500" in plain  # итог: 10000 + 40000 + 7500


def test_generate_batches_and_numbers(tmp_path: Path) -> None:
    """Батчинг по batch_size и нумерация заявок с number_start."""
    if not TEMPLATE.is_file():
        pytest.skip("Эталонный шаблон отсутствует")
    cands = [
        Candidate(last_name=f"Фамилия{i}", first_name="Имя",
                  services=[Service(name="Услуга", hours=1, hour_rate=100.0, cost=100.0)])
        for i in range(3)
    ]
    results = generate(cands, template_path=TEMPLATE, output_dir=tmp_path,
                       number_start=106, batch_size=2)
    assert [p.name for p in results] == ["Заявка 106.docx", "Заявка 107.docx"]
    assert all(p.is_file() for p in results)


def test_main_cli_end_to_end(tmp_path: Path, monkeypatch) -> None:
    """CLI-прогон: xlsx в input/ -> docx в output/ (как это делает run.bat)."""
    if not TEMPLATE.is_file():
        pytest.skip("Эталонный шаблон отсутствует")
    src_dir = Path(__file__).resolve().parents[1]
    inp = tmp_path / "input"
    inp.mkdir()
    wb = Workbook()
    ws = wb.active
    ws.append([None, "Наименование Услуг", "ФИО", "Роль",
               "Объем работ, ч/час", "Стоимость часа, руб.*", "Стоимость услуги, руб."])
    ws.append([None, "Провести анализ", "Сидоров С.", "аналитик", 8, 1000, 8000])
    ws.append([None, "Итого:", None, None, None, None, 8000])
    cand_file = inp / "Шаблон заявки кандидата - Сидоров С..xlsx"
    wb.save(cand_file)

    out = tmp_path / "output"
    from src.main import main

    rc = main(["--input", str(inp), "--config", str(src_dir / "config" / "config.yaml"),
               "--output", str(out)])
    assert rc == 0
    produced = list(out.glob("*.docx"))
    assert len(produced) == 1
    assert produced[0].name == "Заявка 106.docx"
