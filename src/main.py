"""CLI-точка входа: чтение config.yaml, парсинг input/, генерация output/."""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

import yaml

from src.generator import generate
from src.parser import parse_input_dir


def load_config(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _parse_date(value: object) -> date | None:
    if isinstance(value, (date, datetime)):
        return value.date() if isinstance(value, datetime) else value
    if isinstance(value, str) and value.strip():
        return datetime.strptime(value.strip(), "%Y-%m-%d").date()
    return None


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Генератор заявок кандидатов из xlsx")
    ap.add_argument("--config", default="config/config.yaml", help="путь к config.yaml")
    ap.add_argument("--input", default=None, help="каталог с xlsx кандидатов")
    ap.add_argument("--output", default=None, help="каталог для результатов")
    ap.add_argument("--template", default=None, help="docx-шаблон с Jinja2")
    args = ap.parse_args(argv)

    cfg = load_config(Path(args.config))
    paths = cfg.get("paths", {})
    app_cfg = cfg.get("application", {})
    rates = cfg.get("rates", {})
    parser_cfg = cfg.get("parser", {})

    input_dir = Path(args.input or paths.get("input_dir", "input"))
    output_dir = Path(args.output or paths.get("output_dir", "output"))
    template_path = Path(args.template or paths.get("template",
                                                    "templates/заявка_шаблон.docx"))

    if not input_dir.is_dir():
        print(f"Ошибка: каталог {input_dir} не найден", file=sys.stderr)
        return 1
    if not template_path.is_file():
        print(f"Ошибка: шаблон {template_path} не найден", file=sys.stderr)
        return 1

    candidates = parse_input_dir(
        input_dir,
        header_row=int(parser_cfg.get("header_row", 1)),
        max_files=int(parser_cfg.get("max_files", 20)),
    )
    if not candidates:
        print("Кандидаты не найдены — выходим.", file=sys.stderr)
        return 0

    results = generate(
        candidates,
        template_path=template_path,
        output_dir=output_dir,
        number_start=int(app_cfg.get("number_start", 106)),
        base_rate=float(rates.get("base_rate", 1.0)),
    )
    for p in results:
        print(f"Создан файл: {p}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
