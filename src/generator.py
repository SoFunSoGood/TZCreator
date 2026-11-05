"""Агрегация кандидатов и рендер заявок docx по шаблону."""
from __future__ import annotations

from datetime import date
from pathlib import Path
from typing import Optional

from docxtpl import DocxTemplate

from src.models import Application, Candidate, Service


def build_application(candidates: list[Candidate], number: int,
                      created: Optional[date] = None,
                      period_start: Optional[date] = None,
                      period_end: Optional[date] = None,
                      base_rate: float = 1.0) -> Application:
    """Собрать заявку из списка кандидатов."""
    return Application(
        number=number,
        created=created or date.today(),
        candidates=list(candidates),
        period_start=period_start,
        period_end=period_end,
        base_rate=base_rate,
    )


def render_application(app: Application, template_path: Path,
                       output_dir: Path) -> Path:
    """Отрендерить одну заявку в docx и вернуть путь к результату."""
    tpl = DocxTemplate(str(template_path))

    def _empty_service():
        return Service(name="", hours=0.0, hour_rate=0.0, cost=0.0)

    rows = []
    for c in app.candidates:
        for s in (c.services or [_empty_service()]):
            rows.append({
                "full_name": c.full_name,
                "last_name": c.last_name,
                "first_name": c.first_name,
                "middle_name": c.middle_name or "",
                "birth_date": c.birth_date.strftime("%d.%m.%Y") if c.birth_date else "",
                "position": c.position or c.role,
                "role": c.role,
                "unit": c.unit,
                # плоские поля для таблиц услуг (по одной строке на услугу)
                "service": s.name,
                "hours": s.hours,
                "hour_rate": round(s.hour_rate * app.base_rate, 2),
                "cost": round(s.cost * app.base_rate, 2),
            })
    context = {
        "number": app.number,
        "title": app.title,
        "created": app.created.strftime("%d.%m.%Y"),
        "period_start": app.period_start.strftime("%d.%m.%Y") if app.period_start else "",
        "period_end": app.period_end.strftime("%d.%m.%Y") if app.period_end else "",
        "base_rate": app.base_rate,
        "candidates": [
            {
                "full_name": c.full_name,
                "role": c.role,
                "position": c.position,
                "unit": c.unit,
                "birth_date": c.birth_date.strftime("%d.%m.%Y") if c.birth_date else "",
                "total_cost": c.total_cost,
            }
            for c in app.candidates
        ],
        "rows": rows,
        "total": round(sum(r["cost"] for r in rows), 2),
    }
    tpl.render(context)
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / f"{app.title}.docx"
    tpl.save(str(out_path))
    return out_path


def generate(candidates: list[Candidate], template_path: Path, output_dir: Path,
             number_start: int = 106, batch_size: int = 20,
             base_rate: float = 1.0) -> list[Path]:
    """Разбить кандидатов на группы, отрендерить каждую заявку. Возвращает пути."""
    apps: list[Application] = []
    for i, num in enumerate(range(number_start, number_start + 10_000)):
        chunk = candidates[i * batch_size:(i + 1) * batch_size]
        if not chunk:
            break
        apps.append(build_application(chunk, number=num, base_rate=base_rate))
    return [render_application(a, template_path, output_dir) for a in apps]
