from __future__ import annotations

import math
import re
from copy import copy
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Set

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.excel_utils import find_col_index_by_header, find_header_row
from app.settings import COL_EQ_01_VALUE, COL_EQ_09_CODE, COL_MF_DATE


_NONWORD_RE = re.compile(r"[^0-9a-zа-яё]+", flags=re.IGNORECASE)

_TOKEN_RE = re.compile(r"[0-9a-zа-яё]+", re.IGNORECASE)
_YEAR_RE = re.compile(r"(19\d{2}|20\d{2}|21\d{2})")
MANUAL_YEAR_COL_NUM = 6
RU_STOPWORDS = {
    'а',
    'агрегат',
    'без',
    'в',
    'деталь',
    'для',
    'до',
    'же',
    'и',
    'из',
    'к',
    'как',
    'комплект',
    'комплекта',
    'ли',
    'на',
    'но',
    'от',
    'по',
    'при',
    'с',
    'то',
    'узел',
    'часть',
    'это',
}
GENERIC_OP_TOKENS = {
    'вода',
    'воды',
    'жидкости',
    'жидкость',
    'масла',
    'масло',
    'смазка',
    'смазки',
    'фильтр',
    'фильтра',
}
RU_STOPWORDS_STRONG = {
    'а',
    'агрегат',
    'без',
    'в',
    'во',
    'вскрытие',
    'деталь',
    'для',
    'до',
    'же',
    'за',
    'замена',
    'заменить',
    'замены',
    'и',
    'из',
    'или',
    'к',
    'как',
    'комплект',
    'комплекта',
    'крышки',
    'ли',
    'между',
    'на',
    'над',
    'не',
    'необходимо',
    'необходимости',
    'ни',
    'но',
    'от',
    'по',
    'под',
    'при',
    'проверить',
    'проверка',
    'с',
    'со',
    'состояния',
    'то',
    'узел',
    'часть',
    'что',
    'это',
}
GENERIC_ROOTS = ('замен', 'комплект', 'провер', 'вскрыт', 'состоя', 'необходим', 'насос', 'конвейер', 'станц', 'ветв', 'секц', 'крышк', 'направля', 'ролик', 'скребк')


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


def read_target_year_from_manual_by_header(ws_manual: Worksheet, scan_rows: int = 300) -> Optional[int]:
    header_row = find_header_row(ws_manual, [COL_MF_DATE], scan_rows=120)
    col_date = find_col_index_by_header(ws_manual, header_row, COL_MF_DATE)
    max_row = min(ws_manual.max_row, header_row + scan_rows)
    for row_idx in range(header_row + 1, max_row + 1):
        value = ws_manual.cell(row_idx, col_date).value
        if value is None:
            continue
        parsed_date = normalize_to_date(value)
        if parsed_date is not None:
            return parsed_date.year
        match = _YEAR_RE.search(str(value))
        if match:
            year = int(match.group(1))
            if 1900 <= year <= 2100:
                return year
    return None


def read_target_year_from_manual_by_column(
    ws_manual: Worksheet, col_num: int = MANUAL_YEAR_COL_NUM, scan_rows: int = 300
) -> Optional[int]:
    max_row = min(ws_manual.max_row, scan_rows)
    for row_idx in range(1, max_row + 1):
        value = ws_manual.cell(row_idx, col_num).value
        if value is None:
            continue
        match = _YEAR_RE.search(str(value).strip())
        if not match:
            continue
        year = int(match.group(1))
        if 1900 <= year <= 2100:
            return year
    return None


def read_target_year_from_manual(ws_manual: Worksheet) -> Optional[int]:
    year = read_target_year_from_manual_by_header(ws_manual)
    if year is not None:
        return year
    return read_target_year_from_manual_by_column(ws_manual)


def strip_leading_code(name: Any) -> str:
    normalized = normalize_text(name)
    if not normalized:
        return ""
    parts = normalized.split(maxsplit=1)
    return parts[1].strip() if len(parts) == 2 else parts[0]


def token_root(token: str, n: int = 6) -> str:
    token = token.strip()
    return token[:n] if len(token) > n else token


def tokens_ru_set(text: str, min_len: int = 4) -> Set[str]:
    normalized = normalize_text(text)
    if not normalized:
        return set()
    result = set()
    for part in _TOKEN_RE.findall(normalized):
        if len(part) < min_len:
            continue
        if part in RU_STOPWORDS:
            continue
        result.add(token_root(part))
    return result


def is_generic_single_part(part_text: str) -> bool:
    normalized = normalize_text(part_text)
    if not normalized:
        return False
    parts = [p for p in _TOKEN_RE.findall(normalized) if len(p) >= 4 and p not in RU_STOPWORDS]
    return len(parts) == 1 and parts[0] in GENERIC_OP_TOKENS


def ru_light_stem(token: str) -> str:
    if len(token) <= 3:
        return token
    suffixes = (
        "ками", "ями", "ка", "ок", "ек", "ового", "евого", "овыми", "евыми", "овых",
        "евых", "ий", "ый", "ой", "ая", "яя", "ое", "ее", "ые", "ие", "ого", "его",
        "ому", "ему", "ым", "им", "ую", "юю", "ах", "ях", "ам", "ям", "ом", "ем",
        "ов", "ев", "ей", "а", "я", "у", "ю", "е", "и", "ы", "о",
    )
    for suffix in suffixes:
        if token.endswith(suffix) and len(token) - len(suffix) >= 3:
            return token[:-len(suffix)]
    return token


def tokenize_strong(text: Any, min_len: int = 2) -> List[str]:
    normalized = normalize_text_strong(text)
    if not normalized:
        return []
    result = []
    for token in normalized.split():
        if len(token) < min_len:
            continue
        if token in RU_STOPWORDS_STRONG:
            continue
        stemmed = ru_light_stem(token)
        if len(stemmed) >= min_len and stemmed not in RU_STOPWORDS_STRONG:
            result.append(stemmed)
    return result


def has_family(tokens: Set[str], family: tuple[str, ...]) -> bool:
    return any(fragment in token for token in tokens for fragment in family)


def family_overlap(tokens_a: Set[str], tokens_b: Set[str], family: tuple[str, ...]) -> bool:
    for fragment in family:
        if any(fragment in token for token in tokens_a) and any(fragment in token for token in tokens_b):
            return True
    return False


def strip_generic(tokens: Set[str]) -> Set[str]:
    return {
        token for token in tokens
        if not any(root in token for root in GENERIC_ROOTS)
    }


def has_specific_match(part_tokens: Set[str], op_tokens: Set[str]) -> bool:
    part_specific = strip_generic(part_tokens)
    op_specific = strip_generic(op_tokens)
    if not part_specific or not op_specific:
        return False
    if part_specific & op_specific:
        return True
    part_tokens_4 = [token for token in part_specific if len(token) >= 4]
    op_tokens_4 = [token for token in op_specific if len(token) >= 4]
    return any(
        pt in ot or ot in pt
        for pt in part_tokens_4
        for ot in op_tokens_4
    )


def has_roll_root(tokens: Set[str]) -> bool:
    return any("вал" in token for token in tokens)


def build_idf_for_candidates(op_tokens_list: List[Set[str]]) -> Dict[str, float]:
    doc_freq = {}
    for token_set in op_tokens_list:
        for token in token_set:
            doc_freq[token] = doc_freq.get(token, 0) + 1
    total_docs = max(len(op_tokens_list), 1)
    return {
        token: math.log((total_docs + 1) / (count + 1)) + 1.0
        for token, count in doc_freq.items()
    }


def weighted_overlap_score(part_tokens: Set[str], op_tokens: Set[str], idf_map: Dict[str, float]) -> float:
    if not part_tokens or not op_tokens:
        return 0.0
    score = 0.0
    op_has_roll = has_roll_root(op_tokens)
    for part_token in part_tokens:
        weight = idf_map.get(part_token, 1.0)
        if part_token in op_tokens:
            score += weight
            continue
        if "вал" in part_token and op_has_roll:
            score += 0.7 * weight
            continue
        if len(part_token) >= 4:
            if any(part_token in op_token or op_token in part_token for op_token in op_tokens):
                score += 0.7 * weight
    return score / math.sqrt(len(part_tokens) + 1e-9)


def build_equipment_code_to_value_map(ws_eq: Worksheet) -> Dict[str, Any]:
    header_row = find_header_row(ws_eq, [COL_EQ_01_VALUE, COL_EQ_09_CODE], scan_rows=120)
    c_val = find_col_index_by_header(ws_eq, header_row, COL_EQ_01_VALUE)
    c_code = find_col_index_by_header(ws_eq, header_row, COL_EQ_09_CODE)
    result = {}
    for row_idx in range(header_row + 1, ws_eq.max_row + 1):
        code = normalize_text(ws_eq.cell(row_idx, c_code).value)
        if not code:
            continue
        if code not in result:
            result[code] = ws_eq.cell(row_idx, c_val).value
    return result


def equip_contains_either_way(forecast_equip: Any, tech_equip: Any) -> bool:
    a = normalize_text(forecast_equip)
    b = normalize_text(tech_equip)
    if not a or not b:
        return False
    text_a = f" {a} "
    text_b = f" {b} "
    return text_a in text_b or text_b in text_a


_HOURS_RE = re.compile(r"(\d+)\s*час", re.IGNORECASE)
_YEARS_RE = re.compile(r"(\d+)\s*(?:лет|года|год)", re.IGNORECASE)


def norm_period_to_days(value: Any) -> Optional[int]:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text_norm = normalize_text(text)
    if text_norm in {"0", "0.0"}:
        return None
    if "полугод" in text_norm:
        return 182
    match_hours = _HOURS_RE.search(text_norm)
    if match_hours:
        hours = int(match_hours.group(1))
        return max(1, int(round(hours / 24.0)))
    if "раз в год" in text_norm:
        return 365
    match_years = _YEARS_RE.search(text_norm)
    if match_years:
        years = int(match_years.group(1))
        return years * 365
    mapping = {
        "раз в 2 года": 730,
        "раз в 4 года": 1460,
        "раз в 5 лет": 1825,
        "раз в 6 лет": 2190,
    }
    return mapping.get(text_norm)


def iter_dates_in_year_from_base(base: date, step_days: int, target_year: int) -> List[date]:
    if step_days <= 0:
        return []
    year_start = date(target_year, 1, 1)
    year_end = date(target_year, 12, 31)
    current = base
    max_steps = 5000
    steps = 0
    while current < year_start and steps < max_steps:
        current += timedelta(days=step_days)
        steps += 1
    steps = 0
    while current > year_end and steps < max_steps:
        current -= timedelta(days=step_days)
        steps += 1
    if not (year_start <= current <= year_end):
        return []
    result = [current]
    backward = current
    steps = 0
    while steps < max_steps:
        backward -= timedelta(days=step_days)
        if backward < year_start:
            break
        result.append(backward)
        steps += 1
    forward = current
    steps = 0
    while steps < max_steps:
        forward += timedelta(days=step_days)
        if forward > year_end:
            break
        result.append(forward)
        steps += 1
    result.sort()
    return result


def merge_norms_with_defaults(
    norms: Optional[List[tuple[str, str, int]]],
    default_norms: Optional[List[tuple[str, str, int]]],
) -> List[tuple[str, str, int]]:
    result = []
    existing_repairs = set()
    if norms:
        for obj_raw, rep_raw, step_days in norms:
            rep_key = normalize_text(rep_raw)
            result.append((obj_raw, rep_raw, step_days))
            existing_repairs.add(rep_key)
    if default_norms:
        for obj_raw, rep_raw, step_days in default_norms:
            rep_key = normalize_text(rep_raw)
            if rep_key not in existing_repairs:
                result.append((obj_raw, rep_raw, step_days))
                existing_repairs.add(rep_key)
    return result


def autosize_columns(
    ws: Worksheet, *, min_width: float = 8, max_width: float = 30,
    padding: int = 2, scan_rows: Optional[int] = 2000,
) -> None:
    max_row = ws.max_row if scan_rows is None else min(ws.max_row, scan_rows)
    col_max_len = {}
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=ws.max_column):
        for cell in row:
            value = cell.value
            if value is None:
                continue
            if isinstance(value, (datetime, date)):
                text = value.strftime("%d.%m.%Y")
            else:
                text = str(value)
            text_len = len(text.replace("\n", " "))
            if text_len > col_max_len.get(cell.column, 0):
                col_max_len[cell.column] = text_len
    for col_idx in range(1, ws.max_column + 1):
        length = col_max_len.get(col_idx, 0)
        width = min(max(length + padding, min_width), max_width)
        ws.column_dimensions[get_column_letter(col_idx)].width = float(width)
