from __future__ import annotations

from typing import Tuple, Dict, List
from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

from app.settings import (
    DELETE_OBJECT_WORDS,
    DROP_REPAIR_SUBSTRINGS,
    COL_FR_OBJECT,
    COL_FR_REPAIR,
    COL_FR_DATE,
    COL_FR_DATE_SRC,
    COL_FR_VALUE,
)
from app.text_utils import normalize_text
from app.excel_utils import find_header_row, find_col_index_by_header


def normalize_to_date(value):
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


def to_int(value):
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    try:
        return int(s)
    except ValueError:
        return None


def trim_spaces_excel_like(value):
    if value is None:
        return None

    s = str(value).strip()
    if not s:
        return None

    return " ".join(s.split())


def process_factual_repairs(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    stats = {
        "deleted_by_object": 0,
        "deleted_by_repair": 0,
        "dates_normalized": 0,
        "dates_filled": 0,
        "numbers_converted": 0,
    }

    header_row = find_header_row(
        ws,
        [COL_FR_OBJECT, COL_FR_REPAIR, COL_FR_DATE, COL_FR_DATE_SRC, COL_FR_VALUE],
    )
    col1 = find_col_index_by_header(ws, header_row, COL_FR_OBJECT)
    col2 = find_col_index_by_header(ws, header_row, COL_FR_REPAIR)
    col3 = find_col_index_by_header(ws, header_row, COL_FR_DATE)
    col4 = find_col_index_by_header(ws, header_row, COL_FR_DATE_SRC)
    col6 = find_col_index_by_header(ws, header_row, COL_FR_VALUE)

    delete_tokens = {normalize_text(w) for w in DELETE_OBJECT_WORDS if str(w).strip()}
    drop_repairs = tuple(DROP_REPAIR_SUBSTRINGS)

    rows_to_delete = []
    for row in range(header_row + 1, ws.max_row + 1):
        cell1 = ws.cell(row, col1)
        cell2 = ws.cell(row, col2)
        cell3 = ws.cell(row, col3)
        cell4 = ws.cell(row, col4)
        cell6 = ws.cell(row, col6)

        d3 = normalize_to_date(cell3.value)
        if d3:
            cell3.value = d3
            stats["dates_normalized"] += 1

        d4 = normalize_to_date(cell4.value)
        if d4:
            cell4.value = d4

        if not cell3.value and cell4.value:
            cell3.value = cell4.value
            stats["dates_filled"] += 1

        clean1 = trim_spaces_excel_like(cell1.value)
        cell1.value = clean1
        txt1 = normalize_text("" if clean1 is None else clean1)
        if not txt1 or any(tok in txt1 for tok in delete_tokens):
            stats["deleted_by_object"] += 1
            rows_to_delete.append(row)
            continue

        txt2 = normalize_text("" if cell2.value is None else str(cell2.value))
        if txt2 and any(s in txt2 for s in drop_repairs):
            stats["deleted_by_repair"] += 1
            rows_to_delete.append(row)
            continue

        num = to_int(cell6.value)
        if num is not None:
            cell6.value = num
            stats["numbers_converted"] += 1

    for r in reversed(rows_to_delete):
        ws.delete_rows(r)

    ws.delete_cols(col4)
    return len(rows_to_delete), stats
