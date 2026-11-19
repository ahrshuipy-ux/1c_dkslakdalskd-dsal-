from __future__ import annotations

from typing import Tuple, Dict, List

from openpyxl.worksheet.worksheet import Worksheet

from app.settings import (
    DELETE_OBJECT_WORDS,
    DROP_REPAIR_SUBSTRINGS,
    COL_TKOR_NAME,
    COL_TKOR_REPAIR,
    COL_TKOR_ACTION,
)
from app.text_utils import normalize_text
from app.excel_utils import find_header_row, find_col_index_by_header


def process_tech_cards_or(ws: Worksheet) -> Tuple[int, Dict[str, int]]:
    stats = {
        "deleted_by_object_words": 0,
        "deleted_by_repair_type": 0,
        "deleted_not_replacement": 0,
    }

    header_row = find_header_row(
        ws,
        [COL_TKOR_NAME, COL_TKOR_REPAIR, COL_TKOR_ACTION],
        scan_rows=120,
    )

    col1 = find_col_index_by_header(ws, header_row, COL_TKOR_NAME)
    col2 = find_col_index_by_header(ws, header_row, COL_TKOR_REPAIR)
    col3 = find_col_index_by_header(ws, header_row, COL_TKOR_ACTION)

    delete_tokens = {normalize_text(w) for w in DELETE_OBJECT_WORDS if str(w).strip()}
    drop_repairs = tuple(DROP_REPAIR_SUBSTRINGS)

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

    i1 = col1 - 1
    kept = []
    deleted = 0
    for r in data_rows:
        raw = r[i1]
        if raw is None or str(raw).strip() == "":
            deleted += 1
            continue

        txt = normalize_text(str(raw))
        if any(tok in txt for tok in delete_tokens):
            deleted += 1
        else:
            kept.append(r)

    stats["deleted_by_object_words"] = deleted
    data_rows = kept

    i2 = col2 - 1
    kept = []
    deleted = 0
    for r in data_rows:
        txt = normalize_text("" if r[i2] is None else str(r[i2]))
        if txt and any(s in txt for s in drop_repairs):
            deleted += 1
        else:
            kept.append(r)
    stats["deleted_by_repair_type"] = deleted
    data_rows = kept

    i3 = col3 - 1
    kept = []
    deleted = 0
    for r in data_rows:
        txt = normalize_text("" if r[i3] is None else str(r[i3]))
        if "замена" not in txt:
            deleted += 1
        else:
            kept.append(r)
    stats["deleted_not_replacement"] = deleted
    data_rows = kept

    after = len(data_rows)
    total_deleted = before - after

    old_last = ws.max_row
    if old_last > header_row:
        ws.delete_rows(header_row + 1, old_last - header_row)

    for r in data_rows:
        ws.append(list(r))

    return total_deleted, stats
