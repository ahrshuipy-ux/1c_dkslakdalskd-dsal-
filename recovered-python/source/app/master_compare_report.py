from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, date
from typing import Any, Dict, List, Optional, Tuple, Set

from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.excel_utils import find_header_row, find_col_index_by_header
from app.settings import (
    COL_FR_DATE,
    COL_FR_OBJECT,
    COL_FR_REPAIR,
    COL_PR_25,
    COL_PR_26_EQUIP_VALUE,
    COL_PR_27_REPAIR_TYPE,
    COL_MF_DATE,
    COL_MF_OBJECT,
    COL_MF_REPAIR,
    MASTER_COMPARE_SHEET_NAME,
    COL_MRC_01_DATE,
    COL_MRC_02_OBJECT,
    COL_MRC_03_REPAIR,
    COL_MRC_04_DAYS_DIFF,
    COL_MRC_05_PR_DATE,
    COL_MRC_06_PR_REPAIR_TYPE,
    COL_MRC_07_HIT,
    COL_MRC_08,
    COL_MRC_09,
    COL_MRC_10,
    COL_MRC_11,
    MASTER_COMPARE_DAY_WINDOW,
    MASTER_COMPARE_HIT_TEXT,
)

HARD_LOGIC_LABEL = "Жесткая логика:"


def _autosize_columns(ws, *, min_width: float = 8, max_width: float = 30, padding: int = 2, scan_rows: int | None = 2000) -> None:
    max_row = ws.max_row if scan_rows is None else min(ws.max_row, scan_rows)
    col_max_len = {}
    for r in range(1, max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is None:
                continue
            if hasattr(val, "strftime"):
                text = val.strftime("%d.%m.%Y")
            else:
                text = str(val)
            text = text.replace("\n", " ")
            length = len(text)
            if length > col_max_len.get(c, 0):
                col_max_len[c] = length

    for c in range(1, ws.max_column + 1):
        length = col_max_len.get(c, 0)
        width = length + padding
        if width < min_width:
            width = min_width
        if width > max_width:
            width = max_width
        ws.column_dimensions[get_column_letter(c)].width = float(width)


def _normalize_to_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    s = str(value).strip()
    if not s:
        return None
    s = s.split()[0]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _strict_text(v: Any) -> str:
    if v is None:
        return ""
    return str(v)


def _object_key(v: Any) -> str:
    if v is None:
        return ""
    return " ".join(str(v).strip().casefold().split())


def _build_global_pairs_forecast_soft(
    fact_rows: List[Tuple[int, date, Any]],
    cands: List['ForecastRec'],
    window_days: int,
) -> Dict[int, Tuple['ForecastRec', int]]:
    candidates = []
    for fact_row, fact_d, _fact_rep in fact_rows:
        for rec in cands:
            delta = (fact_d - rec.pr_date).days
            abs_delta = abs(delta)
            if abs_delta > window_days:
                continue
            candidates.append((abs_delta, rec.pr_date, fact_row, rec.src_row, rec, delta))

    candidates.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    used_fact_rows = set()
    used_pr_rows = set()
    result = {}
    for _abs_delta, _pr_date, fact_row, pr_row, rec, delta in candidates:
        if fact_row in used_fact_rows or pr_row in used_pr_rows:
            continue
        used_fact_rows.add(fact_row)
        used_pr_rows.add(pr_row)
        result[fact_row] = (rec, delta)

    return result


def _build_global_pairs_manual_soft(
    fact_rows: List[Tuple[int, date, Any]],
    cands: List['ManualRec'],
    window_days: int,
) -> Dict[int, Tuple['ManualRec', int]]:
    candidates = []
    for fact_row, fact_d, _fact_rep in fact_rows:
        for rec in cands:
            delta = (fact_d - rec.mf_date).days
            abs_delta = abs(delta)
            if abs_delta > window_days:
                continue
            candidates.append((abs_delta, rec.mf_date, fact_row, rec.src_row, rec, delta))

    candidates.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    used_fact_rows = set()
    used_mf_rows = set()
    result = {}
    for _abs_delta, _mf_date, fact_row, mf_row, rec, delta in candidates:
        if fact_row in used_fact_rows or mf_row in used_mf_rows:
            continue
        used_fact_rows.add(fact_row)
        used_mf_rows.add(mf_row)
        result[fact_row] = (rec, delta)

    return result


def _build_global_pairs_forecast_hard(
    fact_rows: List[Tuple[int, date, Any]],
    cands: List['ForecastRec'],
    window_days: int,
) -> Dict[int, Tuple['ForecastRec', int]]:
    candidates = []
    for fact_row, fact_d, fact_rep in fact_rows:
        allow_any = _is_emergency_or_unplanned(fact_rep)
        fact_rep_key = _strict_text(fact_rep)
        for rec in cands:
            if not allow_any and _strict_text(rec.pr_repair_type) != fact_rep_key:
                continue
            delta = (fact_d - rec.pr_date).days
            abs_delta = abs(delta)
            if abs_delta > window_days:
                continue
            candidates.append((abs_delta, rec.pr_date, fact_row, rec.src_row, rec, delta))

    candidates.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    used_fact_rows = set()
    used_pr_rows = set()
    result = {}
    for _abs_delta, _pr_date, fact_row, pr_row, rec, delta in candidates:
        if fact_row in used_fact_rows or pr_row in used_pr_rows:
            continue
        used_fact_rows.add(fact_row)
        used_pr_rows.add(pr_row)
        result[fact_row] = (rec, delta)

    return result


def _build_global_pairs_manual_hard(
    fact_rows: List[Tuple[int, date, Any]],
    cands: List['ManualRec'],
    window_days: int,
) -> Dict[int, Tuple['ManualRec', int]]:
    candidates = []
    for fact_row, fact_d, fact_rep in fact_rows:
        allow_any = _is_emergency_or_unplanned(fact_rep)
        fact_rep_key = _strict_text(fact_rep)
        for rec in cands:
            if not allow_any and _strict_text(rec.mf_repair) != fact_rep_key:
                continue
            delta = (fact_d - rec.mf_date).days
            abs_delta = abs(delta)
            if abs_delta > window_days:
                continue
            candidates.append((abs_delta, rec.mf_date, fact_row, rec.src_row, rec, delta))

    candidates.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    used_fact_rows = set()
    used_mf_rows = set()
    result = {}
    for _abs_delta, _mf_date, fact_row, mf_row, rec, delta in candidates:
        if fact_row in used_fact_rows or mf_row in used_mf_rows:
            continue
        used_fact_rows.add(fact_row)
        used_mf_rows.add(mf_row)
        result[fact_row] = (rec, delta)

    return result


def _is_emergency_or_unplanned(rep_fact: Any) -> bool:
    s = _strict_text(rep_fact).strip().casefold()
    if not s:
        return False
    return "аварийн" in s or "внепланов" in s


























@dataclass(frozen=True)
class ForecastRec:
    pr_date: date
    pr_repair_type: Any
    src_row: int


def _build_forecast_index(ws_pr: Worksheet) -> Dict[str, List[ForecastRec]]:
    header_row = find_header_row(ws_pr, [COL_PR_25, COL_PR_26_EQUIP_VALUE, COL_PR_27_REPAIR_TYPE], scan_rows=120)
    c25 = find_col_index_by_header(ws_pr, header_row, COL_PR_25)
    c26 = find_col_index_by_header(ws_pr, header_row, COL_PR_26_EQUIP_VALUE)
    c27 = find_col_index_by_header(ws_pr, header_row, COL_PR_27_REPAIR_TYPE)
    idx = {}
    for r in range(header_row + 1, ws_pr.max_row + 1):
        equip_val = ws_pr.cell(r, c26).value
        key = _object_key(equip_val)
        if not key:
            continue
        d = _normalize_to_date(ws_pr.cell(r, c25).value)
        if d is None:
            continue
        rec = ForecastRec(pr_date=d, pr_repair_type=ws_pr.cell(r, c27).value, src_row=r)
        idx.setdefault(key, []).append(rec)

    for k in idx:
        idx[k].sort(key=lambda x: x.pr_date)
    return idx


def _pick_best_match_forecast_by_date(
    fact_d: date,
    cands: List[ForecastRec],
    window_days: int,
    used_rows: Set[int],
) -> Optional[Tuple[ForecastRec, int]]:
    best = None
    best_abs = 10**9
    for rec in cands:
        if rec.src_row in used_rows:
            continue
        delta = (fact_d - rec.pr_date).days
        ad = abs(delta)
        if ad > window_days:
            continue
        if ad < best_abs:
            best_abs = ad
            best = (rec, delta)
        elif ad == best_abs and best is not None:
            if rec.pr_date < best[0].pr_date:
                best = (rec, delta)
    return best


def _pick_best_match_forecast_hard(
    fact_d: date,
    fact_repair: Any,
    cands: List[ForecastRec],
    window_days: int,
    used_rows: Set[int],
    *,
    allow_any_repair: bool,
) -> Optional[Tuple[ForecastRec, int]]:
    fact_rep_key = _strict_text(fact_repair)
    best = None
    best_abs = 10**9
    for rec in cands:
        if rec.src_row in used_rows:
            continue
        if not allow_any_repair:
            if _strict_text(rec.pr_repair_type) != fact_rep_key:
                continue
        delta = (fact_d - rec.pr_date).days
        ad = abs(delta)
        if ad > window_days:
            continue
        if ad < best_abs:
            best_abs = ad
            best = (rec, delta)
        elif ad == best_abs and best is not None:
            if rec.pr_date < best[0].pr_date:
                best = (rec, delta)
    return best





















@dataclass(frozen=True)
class ManualRec:
    mf_date: date
    mf_repair: Any
    src_row: int


def _build_manual_index(ws_mf: Worksheet) -> Dict[str, List[ManualRec]]:
    header_row = find_header_row(ws_mf, [COL_MF_DATE, COL_MF_OBJECT, COL_MF_REPAIR], scan_rows=120)
    c_date = find_col_index_by_header(ws_mf, header_row, COL_MF_DATE)
    c_obj = find_col_index_by_header(ws_mf, header_row, COL_MF_OBJECT)
    c_rep = find_col_index_by_header(ws_mf, header_row, COL_MF_REPAIR)
    idx = {}
    for r in range(header_row + 1, ws_mf.max_row + 1):
        obj_val = ws_mf.cell(r, c_obj).value
        key = _object_key(obj_val)
        if not key:
            continue
        d = _normalize_to_date(ws_mf.cell(r, c_date).value)
        if d is None:
            continue
        rec = ManualRec(mf_date=d, mf_repair=ws_mf.cell(r, c_rep).value, src_row=r)
        idx.setdefault(key, []).append(rec)

    for k in idx:
        idx[k].sort(key=lambda x: x.mf_date)
    return idx


def _pick_best_match_manual_by_date(
    fact_d: date,
    cands: List[ManualRec],
    window_days: int,
    used_rows: Set[int],
) -> Optional[Tuple[ManualRec, int]]:
    best = None
    best_abs = 10**9
    for rec in cands:
        if rec.src_row in used_rows:
            continue
        delta = (fact_d - rec.mf_date).days
        ad = abs(delta)
        if ad > window_days:
            continue
        if ad < best_abs:
            best_abs = ad
            best = (rec, delta)
        elif ad == best_abs and best is not None:
            if rec.mf_date < best[0].mf_date:
                best = (rec, delta)
    return best


def _pick_best_match_manual_hard(
    fact_d: date,
    fact_repair: Any,
    cands: List[ManualRec],
    window_days: int,
    used_rows: Set[int],
    *,
    allow_any_repair: bool,
) -> Optional[Tuple[ManualRec, int]]:
    fact_rep_key = _strict_text(fact_repair)
    best = None
    best_abs = 10**9
    for rec in cands:
        if rec.src_row in used_rows:
            continue
        if not allow_any_repair:
            if _strict_text(rec.mf_repair) != fact_rep_key:
                continue
        delta = (fact_d - rec.mf_date).days
        ad = abs(delta)
        if ad > window_days:
            continue
        if ad < best_abs:
            best_abs = ad
            best = (rec, delta)
        elif ad == best_abs and best is not None:
            if rec.mf_date < best[0].mf_date:
                best = (rec, delta)
    return best


def _ensure_fresh_sheet(wb, sheet_name: str) -> Worksheet:
    if sheet_name in wb.sheetnames:
        ws_old = wb[sheet_name]
        wb.remove(ws_old)
    return wb.create_sheet(title=sheet_name)


def process_master_compare_file(path: str, *, sheet_names: Optional[Dict[str, str]] = None) -> Dict[str, int]:
    stats = {
        "rows_fact_total": 0,
        "rows_fact_with_date": 0,
        "rows_with_object": 0,
        "forecast_index_keys": 0,
        "manual_index_keys": 0,
        "matched_forecast": 0,
        "no_match_forecast": 0,
        "matched_manual": 0,
        "no_match_manual": 0,
        "matched_forecast_hard": 0,
        "no_match_forecast_hard": 0,
        "matched_manual_hard": 0,
        "no_match_manual_hard": 0,
        "rows_skipped_no_object_in_forecast": 0,
        "rows_with_object_in_forecast": 0,
    }
    wb = load_workbook(path)
    if sheet_names is None:
        sheet_names = {}

    sh_pr = sheet_names.get("Прогноз")
    sh_fr = sheet_names.get("Фактические ремонты")
    sh_mf = sheet_names.get("Ручной прогноз")
    if not sh_pr or sh_pr not in wb.sheetnames:
        raise ValueError("Не найден лист 'Прогноз' в MASTER_RESULT (sheet_names не передан или неверен).")
    if not sh_fr or sh_fr not in wb.sheetnames:
        raise ValueError("Не найден лист 'Фактические ремонты' в MASTER_RESULT (sheet_names не передан или неверен).")
    if not sh_mf or sh_mf not in wb.sheetnames:
        raise ValueError("Не найден лист 'Ручной прогноз' в MASTER_RESULT (sheet_names не передан или неверен).")

    ws_pr = wb[sh_pr]
    ws_fr = wb[sh_fr]
    ws_mf = wb[sh_mf]

    pr_idx = _build_forecast_index(ws_pr)
    stats["forecast_index_keys"] = len(pr_idx)
    mf_idx = _build_manual_index(ws_mf)
    stats["manual_index_keys"] = len(mf_idx)

    ws_out = _ensure_fresh_sheet(wb, MASTER_COMPARE_SHEET_NAME)
    headers = [
        COL_MRC_01_DATE,
        COL_MRC_02_OBJECT,
        COL_MRC_03_REPAIR,
        COL_MRC_04_DAYS_DIFF,
        COL_MRC_05_PR_DATE,
        COL_MRC_06_PR_REPAIR_TYPE,
        COL_MRC_07_HIT,
        COL_MRC_08,
        COL_MRC_09,
        COL_MRC_10,
        COL_MRC_11,
        HARD_LOGIC_LABEL,
        COL_MRC_04_DAYS_DIFF,
        COL_MRC_05_PR_DATE,
        COL_MRC_06_PR_REPAIR_TYPE,
        COL_MRC_07_HIT,
        COL_MRC_08,
        COL_MRC_09,
        COL_MRC_10,
        COL_MRC_11,
    ]
    for c, h in enumerate(headers, start=1):
        ws_out.cell(1, c).value = h

    header_row_fr = find_header_row(ws_fr, [COL_FR_DATE, COL_FR_OBJECT, COL_FR_REPAIR], scan_rows=120)
    c_fr_d = find_col_index_by_header(ws_fr, header_row_fr, COL_FR_DATE)
    c_fr_obj = find_col_index_by_header(ws_fr, header_row_fr, COL_FR_OBJECT)
    c_fr_rep = find_col_index_by_header(ws_fr, header_row_fr, COL_FR_REPAIR)
    fact_items = []
    fact_by_object = {}
    for r in range(header_row_fr + 1, ws_fr.max_row + 1):
        stats["rows_fact_total"] += 1
        d_fact = _normalize_to_date(ws_fr.cell(r, c_fr_d).value)
        obj_fact = ws_fr.cell(r, c_fr_obj).value
        rep_fact = ws_fr.cell(r, c_fr_rep).value
        if d_fact is None:
            continue
        stats["rows_fact_with_date"] += 1
        obj_key = _object_key(obj_fact)
        if not obj_key:
            continue
        stats["rows_with_object"] += 1
        if obj_key not in pr_idx:
            stats["rows_skipped_no_object_in_forecast"] += 1
            continue
        stats["rows_with_object_in_forecast"] += 1
        fact_items.append((r, d_fact, obj_fact, rep_fact, obj_key))
        fact_by_object.setdefault(obj_key, []).append((r, d_fact, rep_fact))

    soft_pr_matches = {}
    soft_mf_matches = {}
    hard_pr_matches = {}
    hard_mf_matches = {}
    for obj_key, fact_rows in fact_by_object.items():
        cands_pr = pr_idx.get(obj_key) or []
        cands_mf = mf_idx.get(obj_key) or []
        soft_pr_matches.update(_build_global_pairs_forecast_soft(fact_rows, cands_pr, MASTER_COMPARE_DAY_WINDOW))
        soft_mf_matches.update(_build_global_pairs_manual_soft(fact_rows, cands_mf, MASTER_COMPARE_DAY_WINDOW))
        hard_pr_matches.update(_build_global_pairs_forecast_hard(fact_rows, cands_pr, MASTER_COMPARE_DAY_WINDOW))
        hard_mf_matches.update(_build_global_pairs_manual_hard(fact_rows, cands_mf, MASTER_COMPARE_DAY_WINDOW))

    out_r = 2
    for r, d_fact, obj_fact, rep_fact, obj_key in fact_items:
        ws_out.cell(out_r, 1).value = d_fact
        ws_out.cell(out_r, 2).value = obj_fact
        ws_out.cell(out_r, 3).value = rep_fact

        best_pr = soft_pr_matches.get(r)
        if best_pr is None:
            stats["no_match_forecast"] += 1
        else:
            rec_pr, delta_pr = best_pr
            stats["matched_forecast"] += 1
            ws_out.cell(out_r, 4).value = abs(delta_pr)
            ws_out.cell(out_r, 5).value = rec_pr.pr_date
            ws_out.cell(out_r, 6).value = rec_pr.pr_repair_type
            ws_out.cell(out_r, 7).value = MASTER_COMPARE_HIT_TEXT

        best_mf = soft_mf_matches.get(r)
        if best_mf is None:
            stats["no_match_manual"] += 1
        else:
            rec_mf, delta_mf = best_mf
            stats["matched_manual"] += 1
            ws_out.cell(out_r, 8).value = abs(delta_mf)
            ws_out.cell(out_r, 9).value = rec_mf.mf_date
            ws_out.cell(out_r, 10).value = rec_mf.mf_repair
            ws_out.cell(out_r, 11).value = MASTER_COMPARE_HIT_TEXT

        best_pr_h = hard_pr_matches.get(r)
        if best_pr_h is None:
            stats["no_match_forecast_hard"] += 1
        else:
            rec_pr_h, delta_pr_h = best_pr_h
            stats["matched_forecast_hard"] += 1
            ws_out.cell(out_r, 13).value = abs(delta_pr_h)
            ws_out.cell(out_r, 14).value = rec_pr_h.pr_date
            ws_out.cell(out_r, 15).value = rec_pr_h.pr_repair_type
            ws_out.cell(out_r, 16).value = MASTER_COMPARE_HIT_TEXT

        best_mf_h = hard_mf_matches.get(r)
        if best_mf_h is None:
            stats["no_match_manual_hard"] += 1
        else:
            rec_mf_h, delta_mf_h = best_mf_h
            stats["matched_manual_hard"] += 1
            ws_out.cell(out_r, 17).value = abs(delta_mf_h)
            ws_out.cell(out_r, 18).value = rec_mf_h.mf_date
            ws_out.cell(out_r, 19).value = rec_mf_h.mf_repair
            ws_out.cell(out_r, 20).value = MASTER_COMPARE_HIT_TEXT

        out_r += 1

    _autosize_columns(ws_out)
    wb.save(path)
    return stats
