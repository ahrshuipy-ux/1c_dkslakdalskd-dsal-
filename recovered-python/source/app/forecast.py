from __future__ import annotations

import math
import re
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, DefaultDict, Dict, Iterable, List, Optional, Set, Tuple

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.excel_utils import find_col_index_by_header, find_header_row
from app.settings import (
    COL_EQ_01_VALUE,
    COL_EQ_09_CODE,
    COL_MF_DATE,
    COL_OBJECT,
    COL_PR_08_CODE,
    COL_PR_09_DATE,
    COL_PR_12,
    COL_PR_14,
    COL_PR_15,
    COL_PR_16,
    COL_PR_17_DATE,
    COL_PR_20,
    COL_PR_21,
    COL_PR_22,
    COL_PR_23,
    COL_PR_25,
    COL_PR_26_EQUIP_VALUE,
    COL_PR_27_REPAIR_TYPE,
    COL_PR_28,
    COL_PR_29,
    COL_PR_COL2_FILTER,
    COL_REPAIR_TYPE,
    COL_TC_01_EQUIP,
    COL_TC_02_REPAIR_TYPE,
    COL_TC_06_WORK,
    FORECAST_SHOP_TO_COL2_ALLOWED,
    COL_PR_07,
)

ProcessResult = Tuple[int, Dict[str, int]]
_TOKEN_RE = re.compile(r"[0-9a-zа-яё]+", re.IGNORECASE)
_NONWORD_RE = re.compile(r"[^0-9a-zа-яё]+", flags=re.IGNORECASE)
_HOURS_RE = re.compile(r"(\d+)\s*час", re.IGNORECASE)
_YEARS_RE = re.compile(r"(\d+)\s*(?:лет|года|год)", re.IGNORECASE)
_YEAR_RE = re.compile(r"(19\d{2}|20\d{2}|21\d{2})")
RU_STOPWORDS = {
    'и', 'в', 'на', 'по', 'к', 'из',
    'для', 'с', 'от', 'до', 'при', 'без',
    'а', 'но', 'ли', 'же', 'то', 'это',
    'как', 'узел', 'деталь', 'часть', 'комплект', 'комплекта',
    'агрегат',
}
RU_STOPWORDS_STRONG = {
    'и', 'в', 'во', 'на', 'к', 'с', 'со',
    'от', 'до', 'по', 'из', 'за', 'под', 'над',
    'для', 'при', 'без', 'или', 'а', 'но', 'что',
    'это', 'как', 'же', 'то', 'ли', 'не', 'ни',
    'замена', 'заменить', 'замены', 'проверка', 'проверить', 'состояния', 'вскрытие',
    'крышки', 'необходимости', 'необходимо', 'между', 'узел', 'деталь', 'часть',
    'комплект', 'комплекта', 'агрегат',
}
GENERIC_OP_TOKENS = {
    'вода', 'воды', 'жидкости', 'жидкость', 'масла', 'масло',
    'смазка', 'смазки', 'фильтр', 'фильтра',
}
REPAIR_PRIORITY = {"текущий (малый) ремонт": 100, "средний ремонт": 200}
F_SEAL = ('уплотн', 'проклад', 'сальник', 'манжет', 'кольц', 'набивк')
F_OIL = ('масл',)
F_PUMP = ('насос',)
F_CONV = ('конвейер', 'цеп', 'звездочк', 'скребк', 'направля', 'ролик', 'натяжн', 'станц')
F_ROLL = ('валк', 'валц', 'вал')
F_ROLLER = ('ролик',)
GENERIC_ROOTS = ('замен', 'комплект', 'провер', 'вскрыт', 'состоя', 'необходим', 'насос', 'конвейер', 'станц', 'ветв', 'секц', 'крышк', 'направля', 'ролик', 'скребк')
DEFAULT_EQUIP_NORM_RULES = [

    {
        "name": 'electric_motor',
        "keywords": ('электродвиг', 'эл.двиг', 'эл двиг', 'двигател', 'двигатель'),
        "norms": [
            ('Электродвигатель', 'Капитальный ремонт', 730),
            ('Электродвигатель', 'Средний ремонт', 365),
            ('Электродвигатель', 'Малый ремонт', 180),
        ],
    },

    {
        "name": 'nasos_centrobezhniy',
        "keywords": ('насос', 'центробеж'),
        "require_all": True,
        "norms": [
            ('Насос центробежный', 'Капитальный ремонт', 1440),
            ('Насос центробежный', 'Средний ремонт', 720),
            ('Насос центробежный', 'Текущий (малый) ремонт', 360),
        ],
    },

    {
        "name": 'konv_lentoch_hiroller',
        "keywords": ('конвейер', 'ленточ'),
        "require_all": True,
        "norms": [
            ('Конвейер ленточный', 'Капитальный ремонт', 720),
            ('Конвейер ленточный', 'Средний ремонт', 360),
            ('Конвейер ленточный', 'Текущий (малый) ремонт', 120),
        ],
    },

    {
        "name": 'teploobmen_kozhuh',
        "keywords": ('аппарат', 'теплообмен', 'кожух'),
        "norms": [
            ('Аппарат теплообменчатый кохужоотрубный', 'Капитальный ремонт', 3650),
            ('Аппарат теплообменчатый кохужоотрубный', 'Средний ремонт', 1800),
            ('Аппарат теплообменчатый кохужоотрубный', 'Текущий (малый) ремонт', 360),
        ],
    },

    {
        "name": 'apparat_nasos_centr',
        "keywords": ('агрегат', 'насос', 'центробеж'),
        "norms": [
            ('Агрегат насосный центробежный', 'Капитальный ремонт', 1440),
            ('Агрегат насосный центробежный', 'Средний ремонт', 720),
            ('Агрегат насосный центробежный', 'Текущий (малый) ремонт', 360),
        ],
    },

    {
        "name": 'emkost',
        "keywords": ('емкост',),
        "norms": [
            ('Емкость', 'Капитальный ремонт', 2160),
            ('Емкость', 'Средний ремонт', 720),
            ('Емкость', 'Текущий (малый) ремонт', 720),
        ],
    },

    {
        "name": 'kondensadootvodchik',
        "keywords": ('конденсатоотводчик',),
        "norms": [
            ('конденсатоотводчик', 'Текущий (малый) ремонт', 90),
        ],
    },

    {
        "name": 'suchka',
        "keywords": ('машина', 'сушил'),
        "require_all": True,
        "norms": [
            ('сушилка', 'Текущий (малый) ремонт', 365),
        ],
    },

    {
        "name": 'apparat_nasos_consol',
        "keywords": ('агрегат', 'насос', 'консольн'),
        "require_all": True,
        "norms": [
            ('Агрегат насосный консольный', 'Средний ремонт', 360),
            ('Агрегат насосный консольный', 'Текущий (малый) ремонт', 180),
        ],
    },

    {
        "name": 'ventilyator_centrobezh',
        "keywords": ('вентилятор', 'центробеж'),
        "require_all": True,
        "norms": [
            ('Вентилятор центробежный', 'Текущий (малый) ремонт', 360),
        ],
    },

    {
        "name": 'vintovoy_nasos',
        "keywords": ('насос', 'винтов'),
        "require_all": True,
        "norms": [
            ('Винтовой насос', 'Капитальный ремонт', 3650),
            ('Винтовой насос', 'Средний ремонт', 1825),
            ('Винтовой насос', 'Текущий (малый) ремонт', 360),
        ],
    },

    {
        "name": 'separator_ciklon',
        "keywords": ('сепарато', 'циклон'),
        "require_all": True,
        "norms": [
            ('сепаратор циклонный', 'Капитальный ремонт', 1460),
            ('сепаратор циклонный', 'Средний ремонт', 730),
            ('сепаратор циклонный', 'Текущий (малый) ремонт', 180),
        ],
    },

    {
        "name": 'separator_cntrobezh',
        "keywords": ('сепарато', 'центробеж'),
        "require_all": True,
        "norms": [
            ('сепаратор центробежн', 'Капитальный ремонт', 1460),
            ('сепаратор центробежн', 'Средний ремонт', 730),
            ('сепаратор центробежн', 'Текущий (малый) ремонт', 180),
        ],
    },
]
NORM_COL_OBJECT_HEADER = 'Наименование объекта'
NORM_COL_REPAIR_TYPE_HEADER = 'Вид ремонта'
NORM_PERIOD_COL_NUM = 6
MANUAL_YEAR_COL_NUM = 6


















@dataclass(frozen=True)
class TechRec:
    equip_roots: Set[str]
    equip_text: str
    op_text: str
    op_tokens: Set[str]
    flags: Dict[str, bool]
    rep_val: Any


@dataclass(frozen=True)
class EquipOpsIndex:
    op_token_to_ids: Dict[str, List[int]]
    rec_ids: List[int]


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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().casefold().split())


def get_default_norms_for_equipment(equip_name: Any) -> Optional[List[tuple[str, str, int]]]:
    s = normalize_text(equip_name)
    if not s:
        return None
    for rule in DEFAULT_EQUIP_NORM_RULES:
        keywords = rule["keywords"]
        if rule.get("require_all", False):
            if all(k in s for k in keywords):
                return rule["norms"]
        elif any(k in s for k in keywords):
            return rule["norms"]
    return None


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


def repair_priority(value: Any) -> int:
    normalized = normalize_text(value)
    if not normalized:
        return -1
    return REPAIR_PRIORITY.get(normalized, 0)


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


def build_tech_index(ws_tc: Worksheet) -> tuple[List[TechRec], Dict[str, List[int]]]:
    header_row = find_header_row(ws_tc, [COL_TC_01_EQUIP, COL_TC_02_REPAIR_TYPE, COL_TC_06_WORK], scan_rows=120)
    c_equip = find_col_index_by_header(ws_tc, header_row, COL_TC_01_EQUIP)
    c_repair = find_col_index_by_header(ws_tc, header_row, COL_TC_02_REPAIR_TYPE)
    c_work = find_col_index_by_header(ws_tc, header_row, COL_TC_06_WORK)
    best_map = {}
    for row_idx in range(header_row + 1, ws_tc.max_row + 1):
        equip_raw = ws_tc.cell(row_idx, c_equip).value
        op_raw = ws_tc.cell(row_idx, c_work).value
        rep_val = ws_tc.cell(row_idx, c_repair).value
        if rep_val in (None, ""):
            continue
        equip_text = normalize_text(equip_raw)
        op_text = normalize_text(op_raw)
        if not equip_text or not op_text:
            continue
        equip_roots = tokens_ru_set(equip_text, min_len=4)
        if not equip_roots:
            continue
        op_tokens = set(tokenize_strong(op_text))
        if not op_tokens:
            continue
        flags = {
            "seal": has_family(op_tokens, F_SEAL),
            "oil": has_family(op_tokens, F_OIL),
            "pump": has_family(op_tokens, F_PUMP),
            "conv": has_family(op_tokens, F_CONV),
            "roll": has_family(op_tokens, F_ROLL) or has_roll_root(op_tokens),
            "roller": has_family(op_tokens, F_ROLLER),
        }
        rec = TechRec(
            equip_roots=equip_roots,
            equip_text=equip_text,
            op_text=op_text,
            op_tokens=op_tokens,
            flags=flags,
            rep_val=rep_val,
        )
        key = (equip_text, op_text)
        pr = repair_priority(rep_val)
        prev = best_map.get(key)
        if prev is None or pr > prev[0]:
            best_map[key] = (pr, rec)
    recs = [rec for _, rec in best_map.values()]
    idx = defaultdict(list)
    for rec_id, rec in enumerate(recs):
        for root in rec.equip_roots:
            idx[root].append(rec_id)
    return recs, dict(idx)


def build_ops_index_by_equip_text(recs: List[TechRec]) -> Dict[str, EquipOpsIndex]:
    by_equip = defaultdict(list)
    for rec_id, rec in enumerate(recs):
        by_equip[rec.equip_text].append(rec_id)
    result = {}
    for equip_text, rec_ids in by_equip.items():
        token_map = defaultdict(list)
        for rec_id in rec_ids:
            for token in recs[rec_id].op_tokens:
                token_map[token].append(rec_id)
        result[equip_text] = EquipOpsIndex(op_token_to_ids=dict(token_map), rec_ids=rec_ids)
    return result


def equip_contains_either_way(forecast_equip: str, tech_equip: str) -> bool:
    a = normalize_text(forecast_equip)
    b = normalize_text(tech_equip)
    if not a or not b:
        return False
    text_a = f" {a} "
    text_b = f" {b} "
    return text_a in text_b or text_b in text_a


def pick_best_repair_type(
    equip_name: Any, part_name: str, tech_recs: List[TechRec],
    tech_idx: Dict[str, List[int]], ops_by_equip_text: Dict[str, EquipOpsIndex],
) -> Any:
    if equip_name in (None, "") or not part_name:
        return None
    equip_name_norm = normalize_text(strip_leading_code(equip_name))
    if not equip_name_norm:
        return None
    part_tokens = set(tokenize_strong(part_name))
    if not part_tokens:
        return None
    if len(part_tokens) == 1 and is_generic_single_part(part_name):
        return None
    part_is_seal = has_family(part_tokens, F_SEAL)
    part_is_oil = has_family(part_tokens, F_OIL)
    part_is_pump = has_family(part_tokens, F_PUMP)
    part_is_conv = has_family(part_tokens, F_CONV)
    part_is_roll = has_family(part_tokens, F_ROLL) or has_roll_root(part_tokens)
    part_is_roller = has_family(part_tokens, F_ROLLER)
    equip_roots = tokens_ru_set(equip_name_norm, min_len=4)
    if not equip_roots:
        return None
    candidate_ids = set()
    for root in equip_roots:
        ids = tech_idx.get(root)
        if ids:
            candidate_ids.update(ids)
    if not candidate_ids:
        return None
    matched_equip_texts = set()
    for rec_id in candidate_ids:
        equip_text = tech_recs[rec_id].equip_text
        if equip_text and equip_contains_either_way(equip_name_norm, equip_text):
            matched_equip_texts.add(equip_text)
    if not matched_equip_texts:
        return None
    selected_equip_text = min(matched_equip_texts, key=lambda text: abs(len(text) - len(equip_name_norm)))
    eq_ops = ops_by_equip_text.get(selected_equip_text)
    if not eq_ops:
        return None
    strong_tokens = [token for token in part_tokens if not any(root in token for root in GENERIC_ROOTS)]
    if not strong_tokens:
        strong_tokens = list(part_tokens)
    strong_tokens.sort(key=len, reverse=True)
    strong_tokens = strong_tokens[:3]
    op_candidate_ids = set()
    for token in strong_tokens:
        ids = eq_ops.op_token_to_ids.get(token)
        if ids:
            op_candidate_ids.update(ids)
    if not op_candidate_ids:
        op_candidate_ids = set(eq_ops.rec_ids)
    idf_map = build_idf_for_candidates([tech_recs[rec_id].op_tokens for rec_id in op_candidate_ids])
    best_val = None
    best_score = float("-inf")
    for rec_id in op_candidate_ids:
        rec = tech_recs[rec_id]
        if rec.flags.get("roll", False) and not part_is_roll:
            continue
        score = weighted_overlap_score(part_tokens, rec.op_tokens, idf_map)
        if score <= 0:
            continue
        specific_ok = has_specific_match(part_tokens, rec.op_tokens)
        allowed = False
        if rec.flags.get("roll", False) and part_is_roll:
            allowed = True
        elif rec.flags.get("roller", False) and part_is_roller and family_overlap(part_tokens, rec.op_tokens, F_ROLLER):
            allowed = True
        elif specific_ok:
            allowed = True
        else:
            if rec.flags.get("seal", False) and part_is_seal and family_overlap(part_tokens, rec.op_tokens, F_SEAL):
                allowed = True
            if rec.flags.get("oil", False) and part_is_oil and family_overlap(part_tokens, rec.op_tokens, F_OIL):
                allowed = True
        if not allowed:
            continue
        if rec.flags.get("seal", False) and not part_is_seal:
            score *= 0.15
        if rec.flags.get("oil", False) and not part_is_oil:
            score *= 0.1
        if part_is_pump and rec.flags.get("pump", False):
            score += 0.2
        if part_is_conv and rec.flags.get("conv", False):
            score += 0.2
        if part_is_pump and rec.flags.get("conv", False) and not rec.flags.get("pump", False):
            score -= 1.0
        if part_is_conv and rec.flags.get("pump", False) and not rec.flags.get("conv", False):
            score -= 1.0
        if rec.flags.get("roll", False) and part_is_roll:
            score += 0.15
        if part_is_roller and rec.flags.get("roller", False):
            score += 0.35
        if score > best_score:
            best_score = score
            best_val = rec.rep_val
    return best_val


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


def build_normatives_index(ws_norm: Worksheet) -> Dict[str, List[tuple[str, str, int]]]:
    header_row = find_header_row(ws_norm, [COL_OBJECT, COL_REPAIR_TYPE], scan_rows=120)
    c_obj = find_col_index_by_header(ws_norm, header_row, COL_OBJECT)
    c_rep = find_col_index_by_header(ws_norm, header_row, COL_REPAIR_TYPE)
    c_per = NORM_PERIOD_COL_NUM
    idx = defaultdict(list)
    for row_idx in range(header_row + 1, ws_norm.max_row + 1):
        obj_raw = ws_norm.cell(row_idx, c_obj).value
        rep_raw = ws_norm.cell(row_idx, c_rep).value
        per_raw = ws_norm.cell(row_idx, c_per).value
        obj_str = "" if obj_raw is None else str(obj_raw).strip()
        rep_str = "" if rep_raw is None else str(rep_raw).strip()
        if not obj_str or not rep_str:
            continue
        rep_norm = normalize_text(rep_str)
        if rep_norm not in ("средний ремонт", "текущий (малый) ремонт"):
            continue
        days = norm_period_to_days(per_raw)
        if days is None:
            continue
        obj_norm = normalize_text(obj_str)
        idx[obj_norm].append((obj_str, rep_str, days))
    return dict(idx)


def norm_equipment_model_tokens(name: str) -> tuple[str, ...]:
    tokens = _TOKEN_RE.findall(name.casefold().replace("ё", "е"))
    for idx, (equipment_type, manufacturer) in enumerate(zip(tokens, tokens[1:])):
        if re.fullmatch(r"[а-я]+", equipment_type) and re.fullmatch(r"[a-z]{2,}", manufacturer):
            return tuple(tokens[idx:])
    return ()


def build_normatives_model_index(
    norm_idx: Dict[str, List[tuple[str, str, int]]],
) -> Dict[tuple[str, ...], List[tuple[tuple[str, ...], List[tuple[str, str, int]]]]]:
    by_model = defaultdict(list)
    for name, norms in norm_idx.items():
        tokens = norm_equipment_model_tokens(name)
        if tokens:
            by_model[tokens[:2]].append((tokens, norms))
    return dict(by_model)


def find_norms_by_model(
    equipment_name: str,
    by_model: Dict[tuple[str, ...], List[tuple[tuple[str, ...], List[tuple[str, str, int]]]]],
) -> List[tuple[str, str, int]]:
    tokens = norm_equipment_model_tokens(equipment_name)
    best_length = 0
    best_norms = []
    for candidate, norms in by_model.get(tokens[:2], ()):
        prefix_length = 0
        for token, other in zip(tokens, candidate):
            if token != other:
                break
            prefix_length += 1
        if prefix_length > best_length:
            best_length = prefix_length
            best_norms = []
        if prefix_length == best_length:
            best_norms.extend(norms)

    # Equally similar models must agree on the period for each repair type.
    by_repair = defaultdict(dict)
    for name, repair, days in sorted(best_norms):
        by_repair[normalize_text(repair)].setdefault(days, (name, repair, days))
    return [next(iter(periods.values())) for periods in by_repair.values() if len(periods) == 1]


def iter_dates_in_year_from_base(base: date, step_days: int, target_year: int) -> Iterable[date]:
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


def append_rows_from_normatives(
    ws_forecast: Worksheet, header_row: int, out25: int, out26: int, out27: int,
    ws_manual: Worksheet, ws_norm: Worksheet, stats: Dict[str, int],
) -> None:
    target_year = read_target_year_from_manual(ws_manual)
    stats["manual_target_year_found"] = 1 if target_year else 0
    if not target_year:
        return
    norm_idx = build_normatives_index(ws_norm)
    stats["norm_index_objects"] = len(norm_idx)
    base_repairs_allowed = {"средний ремонт", "текущий (малый) ремонт", "капитальный ремонт"}
    existing_keys = set()
    for row_idx in range(header_row + 1, ws_forecast.max_row + 1):
        rep_existing = normalize_text(ws_forecast.cell(row_idx, out27).value)
        dt_existing = normalize_to_date(ws_forecast.cell(row_idx, out25).value)
        equip_existing_raw = ws_forecast.cell(row_idx, out26).value
        equip_existing = normalize_text(equip_existing_raw)
        if equip_existing and rep_existing and dt_existing is not None:
            existing_keys.add((equip_existing, rep_existing, dt_existing))
    base_rows = []
    for row_idx in range(header_row + 1, ws_forecast.max_row + 1):
        rep = normalize_text(ws_forecast.cell(row_idx, out27).value)
        if rep not in base_repairs_allowed:
            continue
        planned_date = normalize_to_date(ws_forecast.cell(row_idx, out25).value)
        if planned_date is None:
            continue
        equip_raw = ws_forecast.cell(row_idx, out26).value
        equip_str = "" if equip_raw is None else str(equip_raw).strip()
        if not equip_str:
            continue
        base_rows.append((equip_str, normalize_text(equip_str), planned_date, rep))
    stats["post_base_capital_rows"] = sum(1 for _, _, _, rep in base_rows if rep == "капитальный ремонт")
    stats["post_base_medium_rows"] = sum(1 for _, _, _, rep in base_rows if rep == "средний ремонт")
    stats["post_base_small_rows"] = sum(1 for _, _, _, rep in base_rows if rep == "текущий (малый) ремонт")
    if not base_rows:
        return
    template_row = ws_forecast.max_row
    template_25 = ws_forecast.cell(template_row, out25)
    template_26 = ws_forecast.cell(template_row, out26)
    template_27 = ws_forecast.cell(template_row, out27)
    norms_by_model = build_normatives_model_index(norm_idx)
    seen = set()
    best_by_month = {}
    for equip_raw_to_write, equip_norm, base_date, source_repair in base_rows:
        norms = norm_idx.get(equip_norm)
        if norms is None:
            norms = find_norms_by_model(equip_norm, norms_by_model)
        default_norms = get_default_norms_for_equipment(equip_norm)
        merged_norms = merge_norms_with_defaults(norms, default_norms)
        if norms:
            norm_repairs = {normalize_text(rep_raw) for _, rep_raw, _ in norms}
        else:
            norm_repairs = set()
        if default_norms:
            default_repairs = {normalize_text(rep_raw) for _, rep_raw, _ in default_norms}
        else:
            default_repairs = set()
        missing_repairs_filled = default_repairs - norm_repairs
        if missing_repairs_filled:
            stats["post_default_norms_used"] += 1
            stats["post_default_norm_repairs_added"] += len(missing_repairs_filled)
        if not merged_norms:
            stats["post_skipped_no_norms"] += 1
            continue
        if source_repair == "капитальный ремонт":
            norms_to_apply = merged_norms
        else:
            norms_to_apply = [
                (obj_raw, rep_raw_to_write, step_days)
                for obj_raw, rep_raw_to_write, step_days in merged_norms
                if normalize_text(rep_raw_to_write) == source_repair
            ]
        if not norms_to_apply:
            continue
        for _, rep_raw_to_write, step_days in norms_to_apply:
            rep_norm_key = normalize_text(rep_raw_to_write)
            priority = REPAIR_PRIORITY.get(rep_norm_key, 0)
            for next_date in iter_dates_in_year_from_base(base_date, step_days, target_year):
                exact_key = (equip_norm, rep_norm_key, next_date)
                if exact_key in existing_keys:
                    continue
                if exact_key in seen:
                    continue
                seen.add(exact_key)
                month_key = (equip_norm, next_date.year, next_date.month)
                prev = best_by_month.get(month_key)
                if prev is None:
                    best_by_month[month_key] = (priority, equip_raw_to_write, rep_raw_to_write, next_date)
                    continue
                prev_priority, _, _, prev_date = prev
                if priority > prev_priority or (priority == prev_priority and next_date < prev_date):
                    best_by_month[month_key] = (priority, equip_raw_to_write, rep_raw_to_write, next_date)
    new_rows = [
        (equip_raw_to_write, rep_raw_to_write, next_date)
        for _, equip_raw_to_write, rep_raw_to_write, next_date in best_by_month.values()
    ]
    new_rows.sort(key=lambda item: item[2])
    stats["post_generated_rows"] = len(new_rows)
    if not new_rows:
        return
    write_row = ws_forecast.max_row + 1
    for equip_raw_to_write, rep_raw_to_write, next_date in new_rows:
        c26 = ws_forecast.cell(write_row, out26)
        c27 = ws_forecast.cell(write_row, out27)
        c25 = ws_forecast.cell(write_row, out25)
        c26.value = equip_raw_to_write
        c27.value = rep_raw_to_write
        safe_set_cell_date(ws_forecast, write_row, out25, next_date)
        copy_cell_style(template_26, c26)
        copy_cell_style(template_27, c27)
        copy_cell_style(template_25, c25)
        write_row += 1


def process_forecast(
    ws: Worksheet, chosen_shop: str, ws_equipment: Worksheet,
    ws_tech_cards: Worksheet, ws_manual_forecast: Worksheet,
    ws_normatives: Worksheet,
) -> ProcessResult:
    stats = {
        "tables_stripped": 0,
        "filtered_by_col2": 0,
        "col2_filter_column_missing": 0,
        "col2_filter_skipped_no_mapping": 0,
        "dates_norm_col9": 0,
        "dates_norm_col17": 0,
        "rounded_14": 0,
        "rounded_15": 0,
        "calc_20": 0,
        "calc_21": 0,
        "calc_22": 0,
        "calc_23": 0,
        "calc_25": 0,
        "calc_28": 0,
        "calc_29": 0,
        "equip_map_size": 0,
        "filled_26": 0,
        "skipped_26_no_code": 0,
        "skipped_26_no_code_match": 0,
        "tech_cards_count": 0,
        "tech_index_roots": 0,
        "filled_27_capital": 0,
        "filled_27_from_tech": 0,
        "filled_27_default_medium": 0,
        "skipped_27_no_equip_26": 0,
        "skipped_27_no_match": 0,
        "skipped_27_part_empty": 0,
        "manual_target_year_found": 0,
        "norm_index_objects": 0,
        "post_base_capital_rows": 0,
        "post_generated_rows": 0,
        "post_default_norms_used": 0,
        "post_default_norm_repairs_added": 0,
        "post_skipped_no_norms": 0,
        "post_base_medium_rows": 0,
        "post_base_small_rows": 0,
    }
    strip_excel_tables(ws)
    stats["tables_stripped"] = 1
    equip_map = build_equipment_code_to_value_map(ws_equipment)
    stats["equip_map_size"] = len(equip_map)
    tech_recs, tech_idx = build_tech_index(ws_tech_cards)
    ops_by_equip_text = build_ops_index_by_equip_text(tech_recs)
    stats["tech_cards_count"] = len(tech_recs)
    stats["tech_index_roots"] = len(tech_idx)
    required_headers = [
        COL_PR_09_DATE, COL_PR_14, COL_PR_15, COL_PR_16,
        COL_PR_17_DATE, COL_PR_12, COL_PR_08_CODE, COL_PR_07,
    ]
    header_row = find_header_row(
        ws,
        required_headers,
        scan_rows=120,
    )
    try:
        c2 = find_col_index_by_header(ws, header_row, COL_PR_COL2_FILTER)
    except Exception:
        c2 = None
    c09 = find_col_index_by_header(ws, header_row, COL_PR_09_DATE)
    c14 = find_col_index_by_header(ws, header_row, COL_PR_14)
    c15 = find_col_index_by_header(ws, header_row, COL_PR_15)
    c16 = find_col_index_by_header(ws, header_row, COL_PR_16)
    c17 = find_col_index_by_header(ws, header_row, COL_PR_17_DATE)
    c12 = find_col_index_by_header(ws, header_row, COL_PR_12)
    c08 = find_col_index_by_header(ws, header_row, COL_PR_08_CODE)
    c07 = find_col_index_by_header(ws, header_row, COL_PR_07)
    c10 = 10
    c11 = 11
    allowed_raw = FORECAST_SHOP_TO_COL2_ALLOWED.get(chosen_shop) or ()
    allowed_set = {normalize_text(value) for value in allowed_raw if str(value).strip()}
    rows_to_delete = []
    if c2 is None:
        stats["col2_filter_column_missing"] = 1
    elif not allowed_set:
        stats["col2_filter_skipped_no_mapping"] = 1
    else:
        matched_rows = 0
        for row_idx in range(header_row + 1, ws.max_row + 1):
            value = normalize_text(ws.cell(row_idx, c2).value)
            if value in allowed_set:
                matched_rows += 1
        if matched_rows == 0:
            stats["col2_filter_skipped_no_mapping"] = 1
        else:
            for row_idx in range(header_row + 1, ws.max_row + 1):
                value = normalize_text(ws.cell(row_idx, c2).value)
                if not value or value not in allowed_set:
                    rows_to_delete.append(row_idx)
            for row_idx in reversed(rows_to_delete):
                ws.delete_rows(row_idx)
    stats["filtered_by_col2"] = len(rows_to_delete)

    out20 = ws.max_column + 1
    out21 = out20 + 1
    out22 = out21 + 1
    out23 = out22 + 1
    out25 = out23 + 1
    out26 = out25 + 1
    out27 = out26 + 1
    out28 = out27 + 1
    out29 = out28 + 1
    ws.cell(header_row, out20).value = COL_PR_20
    ws.cell(header_row, out21).value = COL_PR_21
    ws.cell(header_row, out22).value = COL_PR_22
    ws.cell(header_row, out23).value = COL_PR_23
    ws.cell(header_row, out25).value = COL_PR_25
    ws.cell(header_row, out26).value = COL_PR_26_EQUIP_VALUE
    ws.cell(header_row, out27).value = COL_PR_27_REPAIR_TYPE
    ws.cell(header_row, out28).value = COL_PR_28
    ws.cell(header_row, out29).value = COL_PR_29
    today = date.today()
    cell = ws.cell

    for row_idx in range(header_row + 1, ws.max_row + 1):
        cell_09 = cell(row_idx, c09)
        cell_14 = cell(row_idx, c14)
        cell_15 = cell(row_idx, c15)
        cell_16 = cell(row_idx, c16)
        cell_17 = cell(row_idx, c17)
        cell_07 = cell(row_idx, c07)
        cell_12 = cell(row_idx, c12)
        cell_08 = cell(row_idx, c08)
        cell_10 = cell(row_idx, c10)
        cell_11 = cell(row_idx, c11)
        d9 = normalize_to_date(cell_09.value)
        if d9 is not None and safe_set_cell_date(ws, row_idx, c09, d9):
            stats["dates_norm_col9"] += 1
        base_date = normalize_to_date(cell_17.value)
        if base_date is not None and safe_set_cell_date(ws, row_idx, c17, base_date):
            stats["dates_norm_col17"] += 1
        v14_int = round_to_int(cell_14.value)
        if v14_int is not None:
            cell_14.value = v14_int
            stats["rounded_14"] += 1
        v15_int = round_to_int(cell_15.value)
        if v15_int is not None:
            cell_15.value = v15_int
            stats["rounded_15"] += 1
        v14 = to_float(cell_14.value)
        v15 = to_float(cell_15.value)
        v16 = to_float(cell_16.value)
        v22 = round_to_int(v14 / v16) if v14 is not None and v16 not in (None, 0) else None
        cell(row_idx, out22).value = v22
        if v22 is not None:
            stats["calc_22"] += 1
        v23 = round_to_int(v15 * 0.8 / v22) if v15 is not None and v22 not in (None, 0) else None
        cell(row_idx, out23).value = v23
        if v23 is not None:
            stats["calc_23"] += 1
        v20 = round_to_int(v15 - v14) if v15 is not None and v14 is not None else None
        cell(row_idx, out20).value = v20
        if v20 is not None:
            stats["calc_20"] += 1
        v25 = None
        days = safe_to_days(v23)
        if base_date is not None and days is not None:
            try:
                v25 = base_date + timedelta(days=days)
            except OverflowError:
                v25 = None
        if safe_set_cell_date(ws, row_idx, out25, v25):
            stats["calc_25"] += 1
        v28 = None
        if base_date is not None and v22 is not None:
            days_from_base = (today - base_date).days
            v28 = round_to_int(days_from_base * v22)
        cell(row_idx, out28).value = v28
        if v28 is not None:
            stats["calc_28"] += 1
        v21 = round_to_int(v28 / v15 * 100) if v28 is not None and v15 not in (None, 0) else None
        cell(row_idx, out21).value = v21
        if v21 is not None:
            stats["calc_21"] += 1
        v29 = round_to_int(v15 - v28) if v28 is not None and v15 is not None else None
        cell(row_idx, out29).value = v29
        if v29 is not None:
            stats["calc_29"] += 1

        code = normalize_text(cell_08.value)
        if not code:
            cell(row_idx, out26).value = None
            stats["skipped_26_no_code"] += 1
        else:
            equip_value = equip_map.get(code)
            if equip_value is not None:
                cell(row_idx, out26).value = equip_value
                stats["filled_26"] += 1
            else:
                cell(row_idx, out26).value = None
                stats["skipped_26_no_code_match"] += 1
        v12 = normalize_text(cell_12.value)
        v10 = normalize_text(cell_10.value)
        v11 = normalize_text(cell_11.value)
        v07 = normalize_text(cell_07.value)
        if (v10 and v11 and v10 == v11) or (v07 and v12 and v07 == v12):
            cell(row_idx, out27).value = "Капитальный ремонт"
            stats["filled_27_capital"] += 1
            continue
        equip_name = cell(row_idx, out26).value
        if not equip_name:
            stats["skipped_27_no_equip_26"] += 1
            continue
        if not v12:
            stats["skipped_27_part_empty"] += 1
            continue
        rep_val = pick_best_repair_type(
            equip_name=equip_name,
            part_name=v12,
            tech_recs=tech_recs,
            tech_idx=tech_idx,
            ops_by_equip_text=ops_by_equip_text,
        )
        if rep_val is not None:
            cell(row_idx, out27).value = rep_val
            stats["filled_27_from_tech"] += 1
        else:
            stats["skipped_27_no_match"] += 1

    for row_idx in range(header_row + 1, ws.max_row + 1):
        if cell(row_idx, out27).value in (None, ""):
            cell(row_idx, out27).value = "Средний ремонт"
            stats["filled_27_default_medium"] += 1

    append_rows_from_normatives(
        ws_forecast=ws,
        header_row=header_row,
        out25=out25,
        out26=out26,
        out27=out27,
        ws_manual=ws_manual_forecast,
        ws_norm=ws_normatives,
        stats=stats,
    )
    autosize_columns(ws)
    return stats["filtered_by_col2"], stats
