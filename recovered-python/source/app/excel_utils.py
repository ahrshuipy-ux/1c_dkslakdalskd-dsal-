from __future__ import annotations

from typing import List
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from app.text_utils import normalize_text


def get_row_values(ws: Worksheet, row: int, max_col: int) -> List[str]:
    vals = []
    for c in range(1, max_col + 1):
        v = ws.cell(row=row, column=c).value
        vals.append(normalize_text("" if v is None else str(v)))
    return vals


def find_header_row(ws: Worksheet, required_headers: List[str], scan_rows: int = 80) -> int:
    required = {normalize_text(h) for h in required_headers}
    max_row = min(ws.max_row, scan_rows)
    max_col = min(ws.max_column, 200)

    for r in range(1, max_row + 1):
        row_vals = set(get_row_values(ws, r, max_col))
        if required.issubset(row_vals):
            return r

    raise KeyError(
        "Не удалось автоматически найти строку заголовков. Проверь, что в листе есть колонки: "
        + ", ".join(required_headers)
    )


def find_col_index_by_header(ws: Worksheet, header_row: int, header_name: str) -> int:
    wanted = normalize_text(header_name)
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=col).value
        if normalize_text("" if v is None else str(v)) == wanted:
            return col
    raise KeyError(f"Не найден столбец: '{header_name}' (строка заголовков {header_row})")


def copy_worksheet_values(src_ws: Worksheet, dst_ws: Worksheet) -> None:
    for row in src_ws.iter_rows(values_only=False):
        dst_ws.append([cell.value for cell in row])


def load_workbook_safe(path: Path):
    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return load_workbook(path, data_only=False)
    if ext == ".xls":
        raise ValueError(".xls (старый формат) не поддерживается openpyxl. Нужна конвертация в .xlsx.")
    if ext == ".csv":
        raise ValueError(".csv в этой версии не подключён (могу добавить импорт).")
    raise ValueError(f"Неподдерживаемый формат: {ext}")
