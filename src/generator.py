"""Агрегация кандидатов и рендер заявок docx по шаблону."""
from __future__ import annotations

import copy
import re
from datetime import date
from pathlib import Path
from typing import Optional

from docx.oxml.ns import qn
from docxtpl import DocxTemplate

from src.models import Application, Candidate, Service

_W_T = qn("w:t")
_W_P = qn("w:p")
_W_R = qn("w:r")
_W_TC = qn("w:tc")
_W_TR = qn("w:tr")
_W_TBL = qn("w:tbl")
_XML_SPACE = qn("xml:space")

# Строка-цикл: {% for x in items %}{{ ... }}...{% endfor %} внутри одной <w:tr>
_ROW_LOOP_RE = re.compile(
    r"\{%-?\s*for\s+(\w+)\s+in\s+([\w.\[\]'\"]+)\s*-?%\}(.*?)\{%-?\s*endfor\s*-?%\}",
    re.DOTALL,
)
_TAG_SPLIT_RE = re.compile(r"(\{\{-?.*?-?\}\})")


def _cell_prototype_run(tc):
    """Первый ран ячейки — донор форматирования для новых ранов."""
    for p in tc.findall(_W_P):
        r = p.find(_W_R)
        if r is not None:
            return r
    return None


def _clear_runs(p) -> None:
    for child in list(p):
        if child.tag == _W_R:
            p.remove(child)


def _add_text_run(p, proto_run, text: str, first: bool = False) -> None:
    """Добавить ран с текстом в абзац (first=True — в начало абзаца)."""
    if proto_run is not None:
        run = copy.deepcopy(proto_run)
        for child in list(run):
            if child.tag != qn("w:rPr"):
                run.remove(child)
    else:
        run = p.makeelement(_W_R, {})
    t = run.makeelement(_W_T, {})
    t.text = text
    t.set(_XML_SPACE, "preserve")
    run.append(t)
    if first:
        p.insert(0, run)
    else:
        p.append(run)


def normalize_row_loops(tpl: DocxTemplate) -> int:
    """Подготовить строки-циклы шаблона к корректному рендеру docxtpl.

    docxtpl клонирует таблицу-строку для `{% for %}` только когда каждый тег
    лежит в отдельном абзаце/ране (своём <w:t>). Если все теги строки оказались
    в одном текстовом потоке, docxtpl подставляет ВСЕ итерации в одну ячейку:
    строка раздувается до сотен ячеек, tblGrid рассогласуется, и Word не может
    открыть файл («неопознанная ошибка» в /word/document.xml). Здесь мы
    переставляем теги по ячейкам programmatically, править .docx руками не нужно.

    Возвращает число исправленных строк.
    """
    fixed = 0
    body = tpl.get_docx().element.body
    for tr in body.iter(_W_TR):
        full = "".join(t.text or "" for t in tr.iter(_W_T))
        m = _ROW_LOOP_RE.search(full)
        if not m:
            continue
        var, iterable, inner = m.group(1), m.group(2), m.group(3)
        cells = tr.findall(_W_TC)
        parts = [p for p in _TAG_SPLIT_RE.split(inner) if p.strip()]
        if not cells or len(parts) > len(cells):
            # Структура нестандартная — оставляем как есть
            continue
        cell_parts: list[list[str]] = [[] for _ in cells]
        for i, part in enumerate(parts):
            cell_parts[min(i, len(cells) - 1)].append(part)
        for tc, own in zip(cells, cell_parts):
            proto = _cell_prototype_run(tc)
            paras = tc.findall(_W_P)
            if not paras:
                paras = [tc.makeelement(_W_P, {})]
                tc.append(paras[0])
            base_p = paras[0]
            for p in paras[1:]:
                tc.remove(p)
            _clear_runs(base_p)
            for txt in (own or [""]):
                _add_text_run(base_p, proto, txt)
            # дополнительные выражения (если их больше, чем было абзацев) —
            # в отдельные абзацы той же ячейки
            for extra in own[1:]:
                np = copy.deepcopy(base_p)
                _clear_runs(np)
                _add_text_run(np, proto, extra)
                base_p.addnext(np)
                base_p = np
        # {% for %} — в первый абзац первой непустой ячейки,
        # {% endfor %} — в последний абзац последней непустой ячейки
        first_tc = next((tc for tc, own in zip(cells, cell_parts) if own), None)
        last_tc = next(
            (tc for tc, own in reversed(list(zip(cells, cell_parts))) if own), None
        )
        if first_tc is not None and last_tc is not None:
            fp = first_tc.findall(_W_P)[0]
            lp = last_tc.findall(_W_P)[-1]
            _add_text_run(fp, _cell_prototype_run(first_tc),
                          "{%% for %s in %s %%}" % (var, iterable), first=True)
            _add_text_run(lp, _cell_prototype_run(last_tc), "{% endfor %}")
        fixed += 1
    return fixed


def _validate_rendered(tpl: DocxTemplate) -> None:
    """Проверить целостность отрендеренного docx до сохранения на диск."""
    from lxml import etree

    doc = tpl.get_docx()
    body = doc.element.body
    w_t = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
    leftover = [t.text for t in body.iter(w_t)
                if t.text and ("{{" in t.text or "{%" in t.text)]
    if leftover:
        raise ValueError(
            "Шаблон отрендерен с остатками Jinja-тегов: "
            + ", ".join(repr(x[:40]) for x in leftover[:3])
        )
    for tbl in body.iter(_W_TBL):
        grid = tbl.find(qn("w:tblGrid"))
        ncols = len(grid.findall(qn("w:gridCol"))) if grid is not None else 0
        for tr in tbl.findall(_W_TR):
            ncells = len(tr.findall(_W_TC))
            if ncells != ncols:
                raise ValueError(
                    f"Строка таблицы содержит {ncells} ячеек при {ncols} колонках "
                    "шаблона — Word не откроет такой файл. "
                    "Проверьте разметку цикла {% for %} в шаблоне."
                )


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
    normalize_row_loops(tpl)

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
    _validate_rendered(tpl)
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
