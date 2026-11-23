from __future__ import annotations

import sys
from datetime import date
from pathlib import Path
import threading
import traceback
import queue
from typing import List, Dict, Any, Tuple

import tkinter as tk
from tkinter import messagebox, ttk

from openpyxl import Workbook

from app.settings import MASTER_COMPARE_DAY_WINDOW, SUPPORTED_EXT, SHOP_TO_SECTIONS
from app.text_utils import (
    safe_sheet_title_from_filename,
    is_equipment_list_file,
    is_tech_cards_or_file,
    is_forecast_file,
    is_normatives_file,
    is_manual_forecast_file,
    is_factual_repairs_file,
)
from app.excel_utils import load_workbook_safe, copy_worksheet_values
from app.gui import choose_one_shop_gui
from app.processors import run_processor_for_file
from app.master_compare_report import process_master_compare_file
from app.master_compare_stats_report import process_master_compare_stats_file


def _app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _classify_file(file_path: Path) -> str:
    if is_normatives_file(file_path):
        return "Нормативы"
    if is_manual_forecast_file(file_path):
        return "Ручной прогноз"
    if is_tech_cards_or_file(file_path):
        return "Тех. карты ОР"
    if is_factual_repairs_file(file_path):
        return "Фактические ремонты"
    if is_equipment_list_file(file_path):
        return "Список оборудования"
    if is_forecast_file(file_path):
        return "Прогноз"
    return "Прочее"


def _order_key(file_path: Path) -> int:
    ftype = _classify_file(file_path)
    return {
        "Список оборудования": 0,
        "Тех. карты ОР": 1,
        "Нормативы": 2,
        "Фактические ремонты": 3,
        "Ручной прогноз": 4,
        "Прогноз": 5,
    }.get(ftype, 6)


def _show_info_dialog(parent: tk.Tk, title: str, text: str) -> None:
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("700x520")
    win.resizable(True, True)
    win.transient(parent)
    win.grab_set()

    frame = tk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side="right", fill="y")

    txt = tk.Text(frame, wrap="word", yscrollcommand=scrollbar.set)
    txt.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=txt.yview)

    txt.insert("1.0", text)
    txt.config(state="disabled")

    btn_frame = tk.Frame(win)
    btn_frame.pack(fill="x", padx=10, pady=(0, 10))

    def on_next() -> None:
        win.destroy()

    def on_exit() -> None:
        parent.destroy()
        raise SystemExit(0)

    tk.Button(btn_frame, text="Выйти из приложения", command=on_exit, width=20).pack(side="left")
    tk.Button(btn_frame, text="Далее", command=on_next, width=12).pack(side="right")

    win.lift()
    win.focus_force()
    parent.wait_window(win)


def _create_progress_window(parent: tk.Tk, total_steps: int) -> Tuple[tk.Toplevel, ttk.Progressbar, tk.Label]:
    win = tk.Toplevel(parent)
    win.title("Обработка файлов")
    win.geometry("520x140")
    win.resizable(False, False)
    win.transient(parent)

    lbl = tk.Label(win, text="Подготовка...", anchor="w", justify="left")
    lbl.pack(fill="x", padx=12, pady=(12, 6))

    bar = ttk.Progressbar(win, mode="determinate", maximum=max(total_steps, 1))
    bar.pack(fill="x", padx=12, pady=(0, 10))

    win.lift()
    win.attributes("-topmost", True)
    win.after(250, lambda: win.attributes("-topmost", False))
    win.update_idletasks()
    return win, bar, lbl


def _update_progress(win: tk.Toplevel, bar: ttk.Progressbar, lbl: tk.Label, done: int, total: int, text: str) -> None:
    bar["value"] = done
    pct = int(round(done / max(total, 1) * 100))
    lbl.config(text=f"{pct}% ({done}/{total})\n{text}")
    win.update_idletasks()


def _show_load_log_and_validate(parent: tk.Tk, files: List[Path], base_dir: Path) -> bool:
    by_type = {}
    for file_path in files:
        by_type.setdefault(_classify_file(file_path), []).append(file_path)

    required_types = [
        "Список оборудования",
        "Тех. карты ОР",
        "Нормативы",
        "Ручной прогноз",
        "Прогноз",
        "Фактические ремонты",
    ]
    missing = [t for t in required_types if not by_type.get(t)]

    lines = [f"Папка: {base_dir}", "", "Найденные файлы по типам:"]
    for req in required_types:
        lst = by_type.get(req) or []
        if not lst:
            lines.append(f"✖ {req}: НЕ найдено")
        else:
            lines.append(f"✔ {req}: {len(lst)}")
            for fp in lst:
                lines.append(f"   - {fp.name}")

    lines.append("")
    if missing:
        lines.append("Ошибка: отсутствуют обязательные файлы:")
        for m in missing:
            lines.append(f" - {m}")
        messagebox.showerror("Проверка входных файлов", "\n".join(lines), parent=parent)
        return False

    _show_info_dialog(parent, "Проверка входных файлов", "\n".join(lines))
    return True


class ProcessingResult:
    def __init__(self, out_path: Path, summary_lines: List[str], errors_lines: List[str]) -> None:
        self.out_path = out_path
        self.summary_lines = summary_lines
        self.errors_lines = errors_lines


def _process_all_files_worker(
    selected: List[Path],
    chosen_shop: str,
    base_dir: Path,
    q: "queue.Queue[tuple]",
    *,
    window_days: int = MASTER_COMPARE_DAY_WINDOW,
) -> None:
    ctx = {"master_sheet_names": {}}
    master = Workbook()
    master.remove(master.active)

    summary_lines = []
    errors_lines = []

    file_steps = len(selected)
    total_steps = file_steps + 2
    done = 0

    try:
        for file_path in selected:
            done += 1
            q.put(("progress", done, total_steps, file_path.name))

            ext = file_path.suffix.lower()
            sheet_title = safe_sheet_title_from_filename(file_path.name)

            if ext == ".xls":
                errors_lines.append(f"{file_path.name}: .xls не поддерживается.")
                continue
            if ext == ".csv":
                errors_lines.append(f"{file_path.name}: .csv пока не подключён.")
                continue

            try:
                wb = load_workbook_safe(file_path)
            except Exception as exc:
                errors_lines.append(f"{file_path.name}: не удалось открыть: {exc}")
                continue

            src_ws = wb.worksheets[0]

            try:
                matched, proc_name, (total_deleted, stats) = run_processor_for_file(
                    path=file_path,
                    ws=src_ws,
                    chosen_shop=chosen_shop,
                    ctx=ctx,
                )
                if matched:
                    summary_lines.append(f"{file_path.name}: {proc_name} → удалено {total_deleted}")
                    ctx["master_sheet_names"][proc_name] = sheet_title

                    dst = master.create_sheet(title=sheet_title)
                    copy_worksheet_values(src_ws, dst)
                else:
                    summary_lines.append(f"{file_path.name}: пропущен (не используется)")

                if matched and proc_name == "Прогноз":
                    if stats.get("error_no_equipment_list_loaded"):
                        errors_lines.append(f"{file_path.name}: Прогноз, но 'Список оборудования' не загружен до него.")
                    if stats.get("error_no_tech_cards_loaded"):
                        errors_lines.append(f"{file_path.name}: Прогноз, но 'Тех. карты ОР' не загружены до него.")
                    if stats.get("error_no_manual_forecast_loaded"):
                        errors_lines.append(f"{file_path.name}: Прогноз, но 'Ручной прогноз' не загружен до него.")
                    if stats.get("error_no_normatives_loaded"):
                        errors_lines.append(f"{file_path.name}: Прогноз, но 'Нормативы' не загружены до него.")
            except Exception as exc:
                errors_lines.append(f"{file_path.name}: ошибка обработки: {exc}")

        out_path = base_dir / "Результат.xlsx"
        master.save(out_path)

        done += 1
        q.put(("progress", done, total_steps, "Сравнение (master_compare_report)..."))
        try:
            stats_cmp = process_master_compare_file(
                str(out_path),
                sheet_names=ctx.get("master_sheet_names") or {},
                window_days=window_days,
            )
            summary_lines.append(
                f"{out_path.name}: master_compare_report → OK (± {window_days} дней, matched_forecast="
                f"{stats_cmp.get('matched_forecast', 0)}, matched_manual="
                f"{stats_cmp.get('matched_manual', 0)})"
            )
        except Exception as exc:
            errors_lines.append(f"{out_path.name}: master_compare_report → ошибка: {exc}")

        done += 1
        q.put(("progress", done, total_steps, "Сводная статистика (master_compare_stats)..."))
        try:
            stats_sum = process_master_compare_stats_file(str(out_path))
            summary_lines.append(
                f"{out_path.name}: master_compare_stats → OK (repairs="
                f"{stats_sum.get('repairs', 0)}, total_fact="
                f"{stats_sum.get('total_fact', 0)})"
            )
        except Exception as exc:
            errors_lines.append(f"{out_path.name}: master_compare_stats → ошибка: {exc}")

        q.put(("done", ProcessingResult(out_path, summary_lines, errors_lines)))

    except Exception:
        q.put(("fatal", traceback.format_exc()))


def main() -> None:
    root = tk.Tk()
    root.title("Excel Processor")
    root.geometry("520x170")
    root.resizable(False, False)

    status = tk.StringVar(value="Нажмите «Старт»")
    tk.Label(root, textvariable=status, justify="left", anchor="w").pack(fill="x", padx=12, pady=(12, 6))

    max_window_days = (date.max - date.min).days
    window_days_text = tk.StringVar(value="30")
    settings_frame = ttk.Frame(root)
    settings_frame.pack(fill="x", padx=12, pady=(0, 8))
    ttk.Label(settings_frame, text="Допуск оценки, ± дней").pack(side="left")
    ttk.Spinbox(
        settings_frame,
        from_=0,
        to=max_window_days,
        increment=1,
        textvariable=window_days_text,
        width=8,
    ).pack(side="left", padx=(8, 0))

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill="x", padx=12, pady=(0, 12))

    start_btn = tk.Button(btn_frame, text="Старт", width=12)
    start_btn.pack(side="left")

    exit_btn = tk.Button(btn_frame, text="Выход", width=12, command=root.destroy)
    exit_btn.pack(side="right")

    def start_flow() -> None:
        try:
            window_days = int(window_days_text.get())
        except ValueError:
            window_days = -1
        if not 0 <= window_days <= max_window_days:
            messagebox.showwarning(
                "Допуск оценки",
                f"Введите целое число дней от 0 до {max_window_days}.",
                parent=root,
            )
            return

        try:
            start_btn.config(state="disabled")
            status.set("Инициализация...")

            base_dir = _app_base_dir()
            all_files = sorted([
                p for p in base_dir.iterdir()
                if p.is_file()
                and p.suffix.lower() in SUPPORTED_EXT
                and not p.name.startswith(("~$", "~-"))
            ])

            if not all_files:
                messagebox.showinfo("Загрузка", f"В папке {base_dir} не найдено файлов Excel/CSV.", parent=root)
                start_btn.config(state="normal")
                status.set("Нажмите «Старт»")
                return

            if not _show_load_log_and_validate(root, all_files, base_dir):
                start_btn.config(state="normal")
                status.set("Нажмите «Старт»")
                return

            chosen_shop = choose_one_shop_gui(root, list(SHOP_TO_SECTIONS.keys()))
            if chosen_shop is None:
                messagebox.showinfo("Отмена", "Обработка отменена пользователем.", parent=root)
                start_btn.config(state="normal")
                status.set("Нажмите «Старт»")
                return

            selected = sorted(all_files, key=_order_key)
            total_steps = len(selected) + 2
            progress_win, progress_bar, progress_lbl = _create_progress_window(root, total_steps)

            q = queue.Queue()
            finished = {"v": False}
            threading.Thread(
                target=_process_all_files_worker,
                args=(selected, chosen_shop, base_dir, q),
                kwargs={"window_days": window_days},
                daemon=True,
            ).start()

            def poll() -> None:
                try:
                    while True:
                        msg = q.get_nowait()
                        kind = msg[0]

                        if kind == "progress":
                            done, total, text = msg[1], msg[2], msg[3]
                            _update_progress(progress_win, progress_bar, progress_lbl, done, total, text)

                        elif kind == "done":
                            finished["v"] = True
                            try:
                                progress_win.destroy()
                            except Exception:
                                pass

                            res = msg[1]
                            text = f"Готово.\n\nРезультат: {res.out_path.name}\n\n"
                            if res.summary_lines:
                                text += "Отчёт:\n- " + "\n- ".join(res.summary_lines) + "\n\n"
                            if res.errors_lines:
                                text += "Ошибки:\n- " + "\n- ".join(res.errors_lines)

                            messagebox.showinfo("Результат", text, parent=root)
                            status.set("Готово. Нажмите «Старт» для повторного запуска.")
                            start_btn.config(state="normal")
                            return

                        elif kind == "fatal":
                            finished["v"] = True
                            try:
                                progress_win.destroy()
                            except Exception:
                                pass
                            messagebox.showerror("Ошибка", msg[1], parent=root)
                            status.set("Ошибка. Нажмите «Старт» для повторного запуска.")
                            start_btn.config(state="normal")
                            return

                except queue.Empty:
                    pass

                if not finished["v"]:
                    root.after(80, poll)

            status.set("Обработка запущена...")
            root.after(80, poll)

        except Exception:
            messagebox.showerror("FATAL", traceback.format_exc(), parent=root)
            start_btn.config(state="normal")
            status.set("Ошибка. Нажмите «Старт»")

    start_btn.config(command=start_flow)
    root.mainloop()


if __name__ == "__main__":
    main()
