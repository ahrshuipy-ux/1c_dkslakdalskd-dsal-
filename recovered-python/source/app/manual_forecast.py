from __future__ import annotations

from typing import Tuple, Dict, List
from datetime import datetime

from tkinter import messagebox
from openpyxl.worksheet.worksheet import Worksheet

from app.settings import (
    DELETE_OBJECT_WORDS,
    DROP_REPAIR_SUBSTRINGS,
    SHOP_TO_SECTIONS,
    COL_MF_OBJECT,
    COL_MF_REPAIR,
    COL_MF_DATE,
    COL_MF_DATE_SOURCE,
    COL_MF_SECTION,
    COL_MF_VALUE,
)
from app.text_utils import normalize_text
from app.excel_utils import find_header_row, find_col_index_by_header


def to_number(value):
    if value is None:
        return None
    s = str(value).strip().replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

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


def process_manual_forecast(ws: Worksheet, chosen_shop: str) -> Tuple[int, Dict[str, int]]:
    stats = {
        "deleted_by_object": 0,
        "deleted_by_repair": 0,
        "filled_from_col4": 0,
        "deleted_not_in_shops": 0,
        "converted_to_number": 0,
        "normalized_dates": 0,
    }

    if chosen_shop is None:
        messagebox.showinfo("Ручной прогноз", "Обработка отменена.")
        return 0, stats

    header_row = find_header_row(
        ws,
        [
            COL_MF_OBJECT,
            COL_MF_REPAIR,
            COL_MF_DATE,
            COL_MF_DATE_SOURCE,
            COL_MF_SECTION,
            COL_MF_VALUE,
        ],
    )
    col1 = find_col_index_by_header(ws, header_row, COL_MF_OBJECT)
    col2 = find_col_index_by_header(ws, header_row, COL_MF_REPAIR)
    col3 = find_col_index_by_header(ws, header_row, COL_MF_DATE)
    col4 = find_col_index_by_header(ws, header_row, COL_MF_DATE_SOURCE)
    col5 = find_col_index_by_header(ws, header_row, COL_MF_SECTION)
    col6 = find_col_index_by_header(ws, header_row, COL_MF_VALUE)

    sections = SHOP_TO_SECTIONS.get(chosen_shop, [])
    sections_norm = [normalize_text(x) for x in sections if str(x).strip()]
    sections_set = set(sections_norm)
    delete_tokens = [normalize_text(w) for w in DELETE_OBJECT_WORDS if str(w).strip()]
    drop_repairs = tuple(DROP_REPAIR_SUBSTRINGS)

    rows_to_delete = []
    for row in range(header_row + 1, ws.max_row + 1):
        cell1 = ws.cell(row, col1)
        cell2 = ws.cell(row, col2)
        cell3 = ws.cell(row, col3)
        cell4 = ws.cell(row, col4)
        cell5 = ws.cell(row, col5)
        cell6 = ws.cell(row, col6)

        d3 = normalize_to_date(cell3.value)
        if d3:
            cell3.value = d3
            stats["normalized_dates"] += 1

        d4 = normalize_to_date(cell4.value)
        if d4:
            cell4.value = d4
            stats["normalized_dates"] += 1

        obj_txt = normalize_text("" if cell1.value is None else str(cell1.value))
        if obj_txt and delete_tokens and any(tok and tok in obj_txt for tok in delete_tokens):
            stats["deleted_by_object"] += 1
            rows_to_delete.append(row)
            continue

        rep_txt = normalize_text("" if cell2.value is None else str(cell2.value))
        if rep_txt and any(s in rep_txt for s in drop_repairs):
            stats["deleted_by_repair"] += 1
            rows_to_delete.append(row)
            continue

        if not cell3.value and cell4.value:
            cell3.value = cell4.value
            stats["filled_from_col4"] += 1

        sec_txt = normalize_text("" if cell5.value is None else str(cell5.value))
        if sec_txt and sections_set and sec_txt not in sections_set:
            stats["deleted_not_in_shops"] += 1
            rows_to_delete.append(row)
            continue

        num = to_number(cell6.value)
        if num is not None:
            cell6.value = num
            stats["converted_to_number"] += 1

    for r in reversed(rows_to_delete):
        ws.delete_rows(r)

    ws.delete_cols(col4)
    total_deleted = len(rows_to_delete)
    return total_deleted, stats
