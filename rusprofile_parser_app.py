#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Desktop UI wrapper around the Rusprofile scrapy runner."""

from __future__ import annotations

import argparse
import csv
import queue
import re
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Rusprofile Parser"
ROOT_DIR = Path(__file__).resolve().parent
RUNNER = ROOT_DIR / "tools" / "rusprofile-scrapper" / "run_rusprofile_local.py"
DEFAULT_OUTPUT_DIR = ROOT_DIR / "results" / "rusprofile"

# Пищевые коды ОКВЭД по умолчанию: производство и оптовая/розничная торговля продуктами.
FOOD_PRESET = "463900 472900 463600 108220 108900 472400 471100 469000 463800 472910"

INTERESTING_LOG_RE = re.compile(
    r"Saved rows|CSV:|JSON:|XLSX:|ERROR|CRITICAL|Crawled \(\d+\)|Spider closed|closespider",
    re.IGNORECASE,
)


def clean_ids(raw: str) -> list[str]:
    return [part for part in re.split(r"[,\s;]+", raw or "") if part.strip().isdigit()]


def venv_python() -> Path:
    candidate = Path(sys.executable).with_name("python.exe")
    if candidate.exists():
        return candidate
    return Path(sys.executable)


class RusprofileParserApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("1180x760")
        self.queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: Optional[threading.Thread] = None
        self.process: Optional[subprocess.Popen] = None
        self.last_output_dir = DEFAULT_OUTPUT_DIR
        self._build_ui()
        self.root.after(150, self._drain_queue)

    def _build_ui(self) -> None:
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(1, weight=1)

        left = ttk.Frame(self.root, padding=12)
        left.grid(row=0, column=0, rowspan=2, sticky="nsew")
        left.columnconfigure(0, weight=1)

        title = ttk.Label(left, text="Rusprofile Parser", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, sticky="w", pady=(0, 12))

        self.ids_var = tk.StringVar(value=FOOD_PRESET)
        self.limit_var = tk.IntVar(value=300)
        self.concurrency_var = tk.IntVar(value=2)
        self.delay_var = tk.DoubleVar(value=2.0)
        self.output_var = tk.StringVar(value=str(DEFAULT_OUTPUT_DIR))

        ttk.Label(left, text="Коды рубрик Rusprofile (ОКВЭД без точек,\nчерез пробел или запятую)", justify="left").grid(row=1, column=0, sticky="w")
        ids_entry = tk.Text(left, height=4, width=36, wrap="word")
        ids_entry.insert("1.0", FOOD_PRESET)
        ids_entry.grid(row=2, column=0, sticky="ew", pady=(2, 4))
        self.ids_entry = ids_entry
        ttk.Button(left, text="Вернуть пищевой набор", command=self._reset_preset).grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(left, text="Лимит компаний").grid(row=4, column=0, sticky="w")
        ttk.Spinbox(left, from_=1, to=10000, textvariable=self.limit_var, width=10).grid(row=5, column=0, sticky="w", pady=(2, 10))

        ttk.Label(left, text="Параллельные запросы").grid(row=6, column=0, sticky="w")
        ttk.Spinbox(left, from_=1, to=8, textvariable=self.concurrency_var, width=10).grid(row=7, column=0, sticky="w", pady=(2, 10))

        ttk.Label(left, text="Пауза между запросами, сек").grid(row=8, column=0, sticky="w")
        ttk.Spinbox(left, from_=0.5, to=10.0, increment=0.5, textvariable=self.delay_var, width=10).grid(row=9, column=0, sticky="w", pady=(2, 10))

        ttk.Label(left, text="Папка результатов").grid(row=10, column=0, sticky="w")
        ttk.Entry(left, textvariable=self.output_var, width=34).grid(row=11, column=0, sticky="ew", pady=(2, 6))
        ttk.Button(left, text="Выбрать папку", command=self._choose_output_dir).grid(row=12, column=0, sticky="ew", pady=(0, 10))

        self.start_button = ttk.Button(left, text="Запустить сбор", command=self._start)
        self.start_button.grid(row=13, column=0, sticky="ew", pady=(4, 4))
        self.stop_button = ttk.Button(left, text="Остановить", command=self._stop, state="disabled")
        self.stop_button.grid(row=14, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(left, text="Открыть папку результатов", command=self._open_output_dir).grid(row=15, column=0, sticky="ew")

        hint = ttk.Label(
            left,
            text="Собирает компании с rusprofile.ru по кодам\nрубрик: название, ИНН, ОГРН, статус, адрес.\nКод рубрики — это ОКВЭД без точек:\n46.39.00 -> 463900.",
            foreground="#335b9f",
            justify="left",
        )
        hint.grid(row=16, column=0, sticky="w", pady=(18, 0))

        right = ttk.Frame(self.root, padding=(0, 12, 12, 12))
        right.grid(row=0, column=1, rowspan=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=3)
        right.rowconfigure(1, weight=2)

        columns = ("name", "inn", "ogrn", "status", "address")
        self.table = ttk.Treeview(right, columns=columns, show="headings", height=16)
        for column, text, width in (
            ("name", "Компания", 240),
            ("inn", "ИНН", 110),
            ("ogrn", "ОГРН", 130),
            ("status", "Статус", 160),
            ("address", "Адрес", 340),
        ):
            self.table.heading(column, text=text)
            self.table.column(column, width=width, anchor="w")
        self.table.grid(row=0, column=0, sticky="nsew")
        table_scroll = ttk.Scrollbar(right, orient="vertical", command=self.table.yview)
        table_scroll.grid(row=0, column=1, sticky="ns")
        self.table.configure(yscrollcommand=table_scroll.set)

        log_frame = ttk.LabelFrame(right, text="Лог")
        log_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=10, bg="#101827", fg="#e6eefb", insertbackground="#e6eefb")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.status_var = tk.StringVar(value="Готов к работе")
        status = ttk.Label(self.root, textvariable=self.status_var, padding=(12, 6), foreground="#1b4f9c")
        status.grid(row=2, column=0, columnspan=2, sticky="ew")

    def _reset_preset(self) -> None:
        self.ids_entry.delete("1.0", "end")
        self.ids_entry.insert("1.0", FOOD_PRESET)

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.output_var.get() or str(DEFAULT_OUTPUT_DIR))
        if selected:
            self.output_var.set(selected)

    def _open_output_dir(self) -> None:
        path = Path(self.output_var.get() or self.last_output_dir)
        path.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer", str(path)])

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, "Сбор уже идет.")
            return
        ids = clean_ids(self.ids_entry.get("1.0", "end"))
        if not ids:
            messagebox.showwarning(APP_TITLE, "Введите хотя бы один код рубрики, например 463900.")
            return
        if not RUNNER.exists():
            messagebox.showerror(APP_TITLE, f"Не найден скрипт парсера:\n{RUNNER}")
            return
        max_items = max(1, int(self.limit_var.get() or 1))
        concurrency = max(1, int(self.concurrency_var.get() or 1))
        delay = max(0.5, float(self.delay_var.get() or 0.5))
        output_dir = Path(self.output_var.get() or DEFAULT_OUTPUT_DIR)
        output_dir.mkdir(parents=True, exist_ok=True)
        self.last_output_dir = output_dir
        for row in self.table.get_children():
            self.table.delete(row)
        self.log_text.delete("1.0", "end")
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.status_var.set("Сбор запущен")

        command = [
            str(venv_python()),
            str(RUNNER),
            *ids,
            "--max-items", str(max_items),
            "--concurrency", str(concurrency),
            "--delay", str(delay),
            "--output-dir", str(output_dir),
        ]

        def worker() -> None:
            try:
                self.queue.put(("log", f"Коды рубрик: {', '.join(ids)}"))
                creationflags = subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0
                self.process = subprocess.Popen(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    cwd=str(RUNNER.parent),
                    creationflags=creationflags,
                )
                assert self.process.stdout is not None
                for line in self.process.stdout:
                    line = line.rstrip()
                    if line and INTERESTING_LOG_RE.search(line):
                        self.queue.put(("log", line))
                code = self.process.wait()
                if code == 0:
                    self.queue.put(("done", output_dir))
                else:
                    self.queue.put(("error", f"Парсер завершился с кодом {code}. Смотрите лог выше."))
            except Exception as exc:
                self.queue.put(("error", str(exc)))
            finally:
                self.process = None

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _stop(self) -> None:
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self._log("Останавливаю сбор...")

    def _load_latest_csv(self, output_dir: Path) -> int:
        files = sorted(output_dir.glob("rusprofile_*.csv"), key=lambda p: p.stat().st_mtime)
        if not files:
            return 0
        count = 0
        with files[-1].open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                self.table.insert(
                    "",
                    "end",
                    values=(
                        row.get("name", ""),
                        row.get("inn", ""),
                        row.get("ogrn", ""),
                        row.get("current_status", ""),
                        row.get("address", ""),
                    ),
                )
                count += 1
                if count >= 2000:
                    break
        return count

    def _drain_queue(self) -> None:
        try:
            while True:
                event, payload = self.queue.get_nowait()
                if event == "log":
                    self._log(str(payload))
                elif event == "done":
                    self.start_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    output_dir = Path(str(payload))
                    shown = self._load_latest_csv(output_dir)
                    self.status_var.set("Готово")
                    self._log(f"Сбор завершен. В таблице показано строк: {shown}")
                    messagebox.showinfo(APP_TITLE, "Сбор завершен. CSV, JSON и Excel сохранены в папку результатов.")
                elif event == "error":
                    self.start_button.configure(state="normal")
                    self.stop_button.configure(state="disabled")
                    self.status_var.set("Ошибка")
                    self._log(f"Ошибка: {payload}")
                    messagebox.showerror(APP_TITLE, str(payload))
        except queue.Empty:
            pass
        self.root.after(150, self._drain_queue)

    def _log(self, message: str) -> None:
        self.log_text.insert("end", f"{time.strftime('%H:%M:%S')} | {message}\n")
        self.log_text.see("end")

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    parser = argparse.ArgumentParser(description="Desktop Rusprofile parser UI")
    parser.parse_args()
    app = RusprofileParserApp()
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
