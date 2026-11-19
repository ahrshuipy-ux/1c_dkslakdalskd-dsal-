from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from app.excel_utils import find_header_row, find_col_index_by_header
from app.settings import (
    MASTER_COMPARE_SHEET_NAME,
    MASTER_COMPARE_HIT_TEXT,
    MASTER_COMPARE_STATS_SHEET_NAME,
    COL_MRC_03_REPAIR,
    COL_MCS_01_REPAIR,
    COL_MCS_02_FACT_CNT,
    COL_MCS_03_MF_SOFT_CNT,
    COL_MCS_04_MF_HARD_CNT,
    COL_MCS_05_PR_SOFT_CNT,
    COL_MCS_06_PR_HARD_CNT,
    COL_MCS_07_MF_SOFT_PCT,
    COL_MCS_08_MF_HARD_PCT,
    COL_MCS_09_PR_SOFT_PCT,
    COL_MCS_10_PR_HARD_PCT,
)

HARD_LOGIC_LABEL = "Жесткая логика:"


def _strict_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _ensure_fresh_sheet(wb, sheet_name: str) -> Worksheet:
    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    return wb.create_sheet(title=sheet_name)


def _autosize_columns(ws: Worksheet, min_width: float = 8, max_width: float = 40, padding: int = 2) -> None:
    col_max = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = str(v).replace("\n", " ")
            col_max[c] = max(col_max.get(c, 0), len(s))

    for c in range(1, ws.max_column + 1):
        w = col_max.get(c, 0) + padding
        if w < min_width:
            w = min_width
        if w > max_width:
            w = max_width
        ws.column_dimensions[get_column_letter(c)].width = float(w)


@dataclass
class RepairAgg:
    fact_cnt: int = 0
    mf_soft_cnt: int = 0
    mf_hard_cnt: int = 0
    pr_soft_cnt: int = 0
    pr_hard_cnt: int = 0


def process_master_compare_stats_file(
    path: str,
    *,
    compare_sheet_name: str = MASTER_COMPARE_SHEET_NAME,
    out_sheet_name: str = MASTER_COMPARE_STATS_SHEET_NAME,
) -> Dict[str, int]:
    wb = load_workbook(path)
    if compare_sheet_name not in wb.sheetnames:
        raise ValueError(f"Не найден лист сравнения: {compare_sheet_name}")

    ws = wb[compare_sheet_name]
    header_row = find_header_row(ws, [COL_MRC_03_REPAIR], scan_rows=10)
    c_fact_rep = find_col_index_by_header(ws, header_row, COL_MRC_03_REPAIR)
    if ws.max_column < 20:
        raise ValueError(
            f"Лист {compare_sheet_name} имеет {ws.max_column} колонок, ожидается минимум 20."
        )

    c_pr_rep_soft, c_pr_hit_soft = 6, 7
    c_mf_rep_soft, c_mf_hit_soft = 10, 11
    c_pr_rep_hard, c_pr_hit_hard = 15, 16
    c_mf_rep_hard, c_mf_hit_hard = 19, 20
    agg = {}
    rows_total = 0
    rows_used = 0
    total_fact = 0
    total_mf_soft = 0
    total_mf_hard = 0
    total_pr_soft = 0
    total_pr_hard = 0

    for r in range(header_row + 1, ws.max_row + 1):
        rows_total += 1
        fact_rep = _strict_text(ws.cell(r, c_fact_rep).value)
        if not fact_rep:
            continue
        rows_used += 1
        total_fact += 1

        a = agg.setdefault(fact_rep, RepairAgg())
        a.fact_cnt += 1
        if _strict_text(ws.cell(r, c_mf_hit_soft).value) == MASTER_COMPARE_HIT_TEXT and _strict_text(ws.cell(r, c_mf_rep_soft).value):
            a.mf_soft_cnt += 1
            total_mf_soft += 1
        if _strict_text(ws.cell(r, c_mf_hit_hard).value) == MASTER_COMPARE_HIT_TEXT and _strict_text(ws.cell(r, c_mf_rep_hard).value):
            a.mf_hard_cnt += 1
            total_mf_hard += 1
        if _strict_text(ws.cell(r, c_pr_hit_soft).value) == MASTER_COMPARE_HIT_TEXT and _strict_text(ws.cell(r, c_pr_rep_soft).value):
            a.pr_soft_cnt += 1
            total_pr_soft += 1
        if _strict_text(ws.cell(r, c_pr_hit_hard).value) == MASTER_COMPARE_HIT_TEXT and _strict_text(ws.cell(r, c_pr_rep_hard).value):
            a.pr_hard_cnt += 1
            total_pr_hard += 1

    ws_out = _ensure_fresh_sheet(wb, out_sheet_name)
    headers = [
        COL_MCS_01_REPAIR,
        COL_MCS_02_FACT_CNT,
        COL_MCS_03_MF_SOFT_CNT,
        COL_MCS_04_MF_HARD_CNT,
        COL_MCS_05_PR_SOFT_CNT,
        COL_MCS_06_PR_HARD_CNT,
        COL_MCS_07_MF_SOFT_PCT,
        COL_MCS_08_MF_HARD_PCT,
        COL_MCS_09_PR_SOFT_PCT,
        COL_MCS_10_PR_HARD_PCT,
    ]
    for c, h in enumerate(headers, start=1):
        ws_out.cell(1, c).value = h
    items = sorted(agg.items(), key=lambda kv: kv[1].fact_cnt, reverse=True)
    out_r = 2
    for rep, a in items:
        ws_out.cell(out_r, 1).value = rep
        ws_out.cell(out_r, 2).value = a.fact_cnt
        ws_out.cell(out_r, 3).value = a.mf_soft_cnt
        ws_out.cell(out_r, 4).value = a.mf_hard_cnt
        ws_out.cell(out_r, 5).value = a.pr_soft_cnt
        ws_out.cell(out_r, 6).value = a.pr_hard_cnt
        ws_out.cell(out_r, 7).value = None
        ws_out.cell(out_r, 8).value = None
        ws_out.cell(out_r, 9).value = None
        ws_out.cell(out_r, 10).value = None
        out_r += 1

    ws_out.cell(out_r, 1).value = "ИТОГО"
    ws_out.cell(out_r, 2).value = total_fact
    ws_out.cell(out_r, 3).value = total_mf_soft
    ws_out.cell(out_r, 4).value = total_mf_hard
    ws_out.cell(out_r, 5).value = total_pr_soft
    ws_out.cell(out_r, 6).value = total_pr_hard
    denom = max(total_fact, 1)
    ws_out.cell(out_r, 7).value = round(total_mf_soft / denom * 100, 2)
    ws_out.cell(out_r, 8).value = round(total_mf_hard / denom * 100, 2)
    ws_out.cell(out_r, 9).value = round(total_pr_soft / denom * 100, 2)
    ws_out.cell(out_r, 10).value = round(total_pr_hard / denom * 100, 2)
    _autosize_columns(ws_out)
    wb.save(path)

    return {
        "rows_total": rows_total,
        "rows_used": rows_used,
        "repairs": len(agg),
        "total_fact": total_fact,
        "total_mf_soft": total_mf_soft,
        "total_mf_hard": total_mf_hard,
        "total_pr_soft": total_pr_soft,
        "total_pr_hard": total_pr_hard,
    }
