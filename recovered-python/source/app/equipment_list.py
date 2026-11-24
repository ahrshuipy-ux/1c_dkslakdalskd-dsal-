from __future__ import annotations

from typing import Tuple, Dict, List
from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

from app.settings import COL_EQ_DATE, COL_EQ_CODE5, COL_EQ_CODE6
from app.excel_utils import delete_selected_rows, find_header_row, find_col_index_by_header


def normalize_to_date(value):
    """Приводит значение к дате (без времени)."""
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    s = str(value).strip()
    if not s:
        return None

    s = s.split()[0]
    formats = ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y")
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    return None


def to_int_compact(value):
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    s = s.replace(" ", "")
    digits = "".join(ch for ch in s if ch.isdigit())
    if not digits:
        return None

    try:
        return int(digits)
    except ValueError:
        return None


def process_equipment_list(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    stats = {
        "deleted_empty_date": 0,
        "dates_normalized": 0,
        "code5_converted": 0,
        "code6_converted": 0,
    }

    header_row = find_header_row(ws, [COL_EQ_DATE, COL_EQ_CODE5, COL_EQ_CODE6], scan_rows=120)
    col3 = find_col_index_by_header(ws, header_row, COL_EQ_DATE)
    col5 = find_col_index_by_header(ws, header_row, COL_EQ_CODE5)
    col6 = find_col_index_by_header(ws, header_row, COL_EQ_CODE6)
    rows_to_delete = []

    for row in range(header_row + 1, ws.max_row + 1):
        c3 = ws.cell(row, col3)
        c5 = ws.cell(row, col5)
        c6 = ws.cell(row, col6)

        d = normalize_to_date(c3.value)
        if not d:
            rows_to_delete.append(row)
            stats["deleted_empty_date"] += 1
            continue

        c3.value = d
        stats["dates_normalized"] += 1

        v5 = to_int_compact(c5.value)
        if v5 is not None:
            c5.value = v5
            stats["code5_converted"] += 1

        v6 = to_int_compact(c6.value)
        if v6 is not None:
            c6.value = v6
            stats["code6_converted"] += 1

    delete_selected_rows(ws, rows_to_delete)

    return len(rows_to_delete), stats
