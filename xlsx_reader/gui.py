from __future__ import annotations

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .excel_processor import process_excel_file


def select_excel_file() -> str | None:
    """
    Open a file dialog to select an Excel file and return its path or None.
    """
    path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    return path or None


def update_progress(progressbar: ttk.Progressbar, value: int, maximum: int) -> None:
    """
    Update a ttk.Progressbar with the given value/maximum.
    """
    progressbar["maximum"] = maximum
    progressbar["value"] = value
    progressbar.update_idletasks()


def process_file_in_background(
    path: str,
    progressbar: ttk.Progressbar,
    output_text: tk.Text,
    root: tk.Tk,
) -> None:
    """
    Run process_excel_file() on a background thread,
    updating the progressbar and output.
    """

    def worker():
        try:
            def cb(done: int, total: int) -> None:
                root.after(0, update_progress, progressbar, done, total)

            results = process_excel_file(path, cb)

            def show_results():
                output_text.configure(state="normal")
                output_text.delete("1.0", "end")
                for sheet, count in results.items():
                    output_text.insert("end", f"{sheet}: {count} rows\n")
                output_text.configure(state="disabled")

            root.after(0, show_results)
        except Exception as exc:  # pragma: no cover
            root.after(0, messagebox.showerror, "Error", str(exc))

    threading.Thread(target=worker, daemon=True).start()


def create_main_window() -> tk.Tk:
    """
    Build the main window and wire up events.
    """
    root = tk.Tk()
    root.title("XLSX Reader")

    frame = ttk.Frame(root, padding=12)
    frame.pack(fill="both", expand=True)

    choose_btn = ttk.Button(frame, text="Select Excel file")
    choose_btn.pack(anchor="w")

    progressbar = ttk.Progressbar(frame, mode="determinate")
    progressbar.pack(fill="x", pady=8)

    output = tk.Text(frame, height=12, width=64)
    output.configure(state="disabled")
    output.pack(fill="both", expand=True)

    def choose_and_run():
        path = select_excel_file()
        if not path:
            return
        update_progress(progressbar, 0, 1)
        process_file_in_background(path, progressbar, output, root)

    choose_btn.configure(command=choose_and_run)
    return root


def run_app() -> None:
    """
    Create and run the app.
    """
    root = create_main_window()
    root.mainloop()
