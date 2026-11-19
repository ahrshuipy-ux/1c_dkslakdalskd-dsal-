from __future__ import annotations

import math
import re
from copy import copy
from datetime import date, datetime
from typing import Any, Optional

from openpyxl.worksheet.worksheet import Worksheet


_NONWORD_RE = re.compile(r"[^0-9a-zа-яё]+", flags=re.IGNORECASE)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().casefold().split())


def normalize_text_strong(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = _NONWORD_RE.sub(" ", text)
    return " ".join(text.split())


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def round_to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return math.floor(number + 0.5) if number >= 0 else math.ceil(number - 0.5)


def normalize_to_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    text = text.split()[0]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def safe_to_days(value: Optional[float], max_abs_days: int = 36500) -> Optional[int]:
    if value is None:
        return None
    try:
        days = int(round(float(value)))
    except (ValueError, TypeError, OverflowError):
        return None
    if abs(days) > max_abs_days:
        return None
    return days


def safe_set_cell_date(ws: Worksheet, row: int, col: int, dt: Optional[date]) -> bool:
    cell = ws.cell(row, col)
    if dt is None:
        cell.value = None
        return False
    try:
        cell.value = dt
        return True
    except ValueError:
        cell.value = None
        return False


def copy_cell_style(src_cell, dst_cell) -> None:
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.protection = copy(src_cell.protection)
    dst_cell.number_format = src_cell.number_format


def strip_excel_tables(ws: Worksheet) -> None:
    try:
        table_names = list(ws.tables.keys())
    except AttributeError:
        return
    for table_name in table_names:
        del ws.tables[table_name]
