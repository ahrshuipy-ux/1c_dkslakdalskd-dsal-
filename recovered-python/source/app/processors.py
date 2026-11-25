from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Dict, Tuple, Any, Optional

from openpyxl.worksheet.worksheet import Worksheet

from app.text_utils import (
    is_normatives_file,
    is_manual_forecast_file,
    is_tech_cards_or_file,
    is_factual_repairs_file,
    is_equipment_list_file,
    is_forecast_file,
)


from app.normatives import process_normatives_sheet
from app.manual_forecast import process_manual_forecast
from app.tech_cards_or import process_tech_cards_or
from app.factual_repairs import process_factual_repairs
from app.equipment_list import process_equipment_list
from app.forecast import process_forecast

ProcessResult = Tuple[int, Dict[str, int]]
Detector = Callable[[Any], bool]


@dataclass(frozen=True)
class Processor:
    detector: Detector
    runner: Callable[..., ProcessResult]
    needs_shop: bool
    name: str


PROCESSORS = [
    Processor(is_normatives_file, process_normatives_sheet, True, "Нормативы"),
    Processor(is_manual_forecast_file, process_manual_forecast, True, "Ручной прогноз"),
    Processor(is_tech_cards_or_file, process_tech_cards_or, False, "Тех. карты ОР"),
    Processor(is_factual_repairs_file, process_factual_repairs, False, "Фактические ремонты"),
    Processor(is_equipment_list_file, process_equipment_list, False, "Список оборудования"),
    Processor(is_forecast_file, process_forecast, True, "Прогноз"),
]


def run_processor_for_file(
    path,
    ws: Worksheet,
    chosen_shop: str,
    ctx: Optional[Dict[str, Any]] = None,
) -> Tuple[bool, str, ProcessResult]:
    if ctx is None:
        ctx = {}

    for p in PROCESSORS:
        if not p.detector(path):
            continue

        if p.name == "Нормативы":
            res = p.runner(ws, chosen_shop)
            ctx["ws_normatives"] = ws
            return True, p.name, res

        if p.name == "Ручной прогноз":
            res = p.runner(ws, chosen_shop)
            ctx["ws_manual_forecast"] = ws
            return True, p.name, res

        if p.name == "Фактические ремонты":
            res = p.runner(ws)
            ctx["ws_factual_repairs"] = ws
            return True, p.name, res

        if p.name == "Список оборудования":
            res = p.runner(ws)
            ctx["ws_equipment"] = ws
            return True, p.name, res

        if p.name == "Тех. карты ОР":
            res = p.runner(ws)
            ctx["ws_tech_cards"] = ws
            return True, p.name, res

        if p.name == "Прогноз":
            ws_equipment = ctx.get("ws_equipment")
            ws_tech_cards = ctx.get("ws_tech_cards")
            ws_manual_forecast = ctx.get("ws_manual_forecast")
            ws_normatives = ctx.get("ws_normatives")

            if ws_equipment is None:
                return True, p.name, (0, {"error_no_equipment_list_loaded": 1})
            if ws_tech_cards is None:
                return True, p.name, (0, {"error_no_tech_cards_loaded": 1})
            if ws_manual_forecast is None:
                return True, p.name, (0, {"error_no_manual_forecast_loaded": 1})
            if ws_normatives is None:
                return True, p.name, (0, {"error_no_normatives_loaded": 1})

            return True, p.name, p.runner(
                ws,
                chosen_shop,
                ws_equipment,
                ws_tech_cards,
                ws_manual_forecast,
                ws_normatives,
                ws_factual_repairs=ctx.get("ws_factual_repairs"),
            )

        if p.needs_shop:
            return True, p.name, p.runner(ws, chosen_shop)
        return True, p.name, p.runner(ws)

    return False, "Без обработки", (0, {})
