"""Структуры данных кандидата и заявки."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass
class Service:
    """Одна услуга (строка объёмов/стоимости) кандидата."""

    name: str = ""                 # наименование услуги
    hours: float = 0.0             # объём работ, ч/час
    hour_rate: float = 0.0         # стоимость часа, руб.
    cost: float = 0.0              # стоимость услуги, руб.


@dataclass
class Candidate:
    """Данные одного кандидата, прочитанные из xlsx."""

    last_name: str = ""
    first_name: str = ""
    middle_name: Optional[str] = None
    birth_date: Optional[date] = None
    position: str = ""
    unit: str = ""
    role: str = ""                       # роль в проекте
    services: list[Service] = field(default_factory=list)
    source_file: str = ""  # откуда прочитан (для отладки)

    @property
    def full_name(self) -> str:
        """ФИО в именительном падеже."""
        parts = [self.last_name, self.first_name, self.middle_name or ""]
        return " ".join(p for p in parts if p)

    @property
    def total_cost(self) -> float:
        """Суммарная стоимость услуг кандидата."""
        return round(sum(s.cost for s in self.services), 2)


@dataclass
class Application:
    """Агрегированная заявка (один docx на группу кандидатов)."""

    number: int
    created: date
    candidates: list[Candidate] = field(default_factory=list)
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    base_rate: float = 1.0

    @property
    def title(self) -> str:
        return f"Заявка {self.number}"
