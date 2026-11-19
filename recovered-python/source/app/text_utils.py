from __future__ import annotations

import re
from pathlib import Path
from typing import List

from app.settings import NORMATIVE_KEYWORDS


def normalize_text(s: str) -> str:
    s = re.sub(r"\s+", " ", s)
    return s.casefold()


def safe_sheet_title_from_filename(filename: str) -> str:
    base = Path(filename).stem
    base = re.sub(r"[:\\/?*\[\]]", "_", base).strip()
    if not base:
        base = "Sheet"
    return base[:31]


def is_normatives_file(path: Path) -> bool:
    name = path.stem.casefold()
    return any(k in name for k in NORMATIVE_KEYWORDS)


def clean_user_tokens(raw: str) -> List[str]:
    if not raw:
        return []
    parts = [p.strip() for p in raw.split(",")]
    cleaned = []
    for p in parts:
        p = p.strip(" \t\r\n\"'\x60“”‘’()[]{}<>")
        p = re.sub(r"[^\w\s\-А-Яа-яЁё]", " ", p, flags=re.UNICODE)
        p = re.sub(r"\s+", " ", p).strip()
        if p:
            cleaned.append(p.casefold())

    out = []
    seen = set()
    for t in cleaned:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out


def is_manual_forecast_file(path) -> bool:
    name = path.stem.casefold()
    return "ручной" in name and "прогноз" in name


def is_tech_cards_or_file(path) -> bool:
    name = path.stem.casefold()
    return ("технолог" in name and "карт" in name and "ор" in name) or (
        "тех" in name and "карт" in name and "ор" in name
    )


def is_factual_repairs_file(path) -> bool:
    name = path.stem.casefold()
    return "фактичес" in name and "ремонт" in name


def is_equipment_list_file(path) -> bool:
    name = path.stem.casefold()
    return "список" in name and "оборуд" in name


def is_forecast_file(path) -> bool:
    name = path.stem.casefold()
    return "прогноз" in name and "ручн" not in name
