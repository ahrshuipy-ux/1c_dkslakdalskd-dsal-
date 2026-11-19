from __future__ import annotations

import tkinter as tk
from tkinter import messagebox
from typing import List, Optional


def choose_one_shop_gui(parent: tk.Tk, shop_names: List[str]) -> Optional[str]:
    if not shop_names:
        messagebox.showerror("Цех", "Список цехов пуст.", parent=parent)
        return None

    win = tk.Toplevel(parent)
    win.title("Выбор цеха")
    win.geometry("520x420")
    win.resizable(True, True)

    win.transient(parent)
    win.grab_set()

    tk.Label(
        win,
        text="Выберите предприятие.\nДанные будут отфильтрованы по цехам выбранного предприятия.",
        justify="left",
        anchor="w",
    ).pack(fill="x", padx=10, pady=10)

    frame = tk.Frame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(frame, selectmode=tk.SINGLE, yscrollcommand=scrollbar.set)
    for shop in shop_names:
        listbox.insert(tk.END, shop)

    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    selected = {"value": None}

    def on_ok() -> None:
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Цех", "Не выбран цех.", parent=win)
            return
        selected["value"] = shop_names[sel[0]]
        win.destroy()

    def on_cancel() -> None:
        win.destroy()

    btn_frame = tk.Frame(win)
    btn_frame.pack(fill="x", padx=10, pady=10)

    tk.Button(btn_frame, text="Далее", command=on_ok, width=12).pack(side="right", padx=5)
    tk.Button(btn_frame, text="Отмена", command=on_cancel, width=12).pack(side="right", padx=5)

    win.lift()
    win.focus_force()

    parent.wait_window(win)
    return selected["value"]
