from __future__ import annotations

import re
from typing import List, Dict, Tuple, Set

from tkinter import messagebox
from openpyxl.worksheet.worksheet import Worksheet

from app.settings import (
    SHOP_TO_SECTIONS,
    DELETE_OBJECT_WORDS,
    COL_STOP_DIV,
    COL_OBJECT,
    COL_REPAIR_TYPE,
    DROP_REPAIR_SUBSTRINGS,
)
from app.text_utils import normalize_text
from app.excel_utils import find_header_row, find_col_index_by_header

_SPLIT_RX = re.compile(r"[;,|/\\]+|\s{2,}", flags=re.UNICODE)
_REPAIR_TRAILING_NUM_RX = re.compile(r"\s+\d+\s*$", flags=re.UNICODE)


def split_to_tokens_exact(cell_value: str) -> List[str]:
    s = normalize_text(cell_value)
    if not s:
        return []
    parts = [p.strip() for p in _SPLIT_RX.split(s)]
    return [p for p in parts if p]


def strip_repair_trailing_number(value: str) -> str:
    s = value.strip()
    return _REPAIR_TRAILING_NUM_RX.sub("", s).strip()


def process_normatives_sheet(ws: Worksheet, chosen_shop: str) -> Tuple[int, Dict[str, int]]:
    stats = {
        "deleted_not_in_shops": 0,
        "deleted_by_object_words": 0,
        "deleted_by_repair_type": 0,
        "repair_type_cleaned": 0,
    }

    if chosen_shop is None:
        messagebox.showinfo("Нормативы", "Обработка файла отменена.")
        return 0, stats

    header_row = find_header_row(
        ws,
        [COL_STOP_DIV, COL_OBJECT, COL_REPAIR_TYPE],
        scan_rows=120,
    )
    col_stop = find_col_index_by_header(ws, header_row, COL_STOP_DIV)
    col_obj = find_col_index_by_header(ws, header_row, COL_OBJECT)
    col_rep = find_col_index_by_header(ws, header_row, COL_REPAIR_TYPE)

    raw_sections = SHOP_TO_SECTIONS.get(chosen_shop, [])
    sections_norm = [normalize_text(x) for x in raw_sections if str(x).strip()]

    if not sections_norm:
        messagebox.showwarning(
            "Нормативы",
            f"Для цеха '{chosen_shop}' список участков пуст.\nФильтрация по участкам пропущена.",
        )

    del_tokens = [
        normalize_text(str(x)) for x in DELETE_OBJECT_WORDS if str(x).strip()
    ]
    del_tokens_set = set(del_tokens)
    max_col = ws.max_column
    data_rows = list(
        ws.iter_rows(
            min_row=header_row + 1,
            max_row=ws.max_row,
            max_col=max_col,
            values_only=True,
        )
    )

    before = len(data_rows)

    if sections_norm:
        sections_set = set(sections_norm)
        stop_idx = col_stop - 1

        def row_section_ok(r: Tuple) -> bool:
            v = r[stop_idx]
            txt = "" if v is None else str(v)
            tokens = split_to_tokens_exact(txt)
            return any(t in sections_set for t in tokens)

        kept = []
        deleted_cnt = 0
        for r in data_rows:
            if row_section_ok(r):
                kept.append(r)
            else:
                deleted_cnt += 1

        stats["deleted_not_in_shops"] = deleted_cnt
        data_rows = kept

    if del_tokens_set:
        obj_idx = col_obj - 1

        def has_bad_token(r: Tuple) -> bool:
            v = r[obj_idx]
            txt = normalize_text("" if v is None else str(v))
            return any(tok in txt for tok in del_tokens_set)

        kept = []
        deleted_cnt = 0
        for r in data_rows:
            if has_bad_token(r):
                deleted_cnt += 1
            else:
                kept.append(r)

        stats["deleted_by_object_words"] = deleted_cnt
        data_rows = kept

    rep_idx = col_rep - 1
    subs = DROP_REPAIR_SUBSTRINGS
    kept = []
    deleted_cnt = 0
    for r in data_rows:
        row = list(r)
        v = row[rep_idx]
        raw = "" if v is None else str(v).strip()

        cleaned_raw = strip_repair_trailing_number(raw)
        if cleaned_raw != raw:
            stats["repair_type_cleaned"] += 1

        row[rep_idx] = cleaned_raw if cleaned_raw else None

        txt = normalize_text(cleaned_raw)
        if txt and any(s in txt for s in subs):
            deleted_cnt += 1
        else:
            kept.append(tuple(row))

    stats["deleted_by_repair_type"] = deleted_cnt
    data_rows = kept

    after = len(data_rows)
    total_deleted = before - after
    old_last = ws.max_row
    if old_last > header_row:
        ws.delete_rows(header_row + 1, old_last - header_row)

    for r in data_rows:
        ws.append(list(r))

    return total_deleted, stats
