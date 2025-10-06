from __future__ import annotations
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Dict
from .excel_processor import process_excel_file

def select_excel_file(entry_widget: tk.Entry) -> None:
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx;*.xls")],
    )
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def update_progress(progressbar: ttk.Progressbar, var: tk.DoubleVar, value: float) -> None:
    var.set(max(0.0, min(1.0, float(value))))
    progressbar.update_idletasks()

def process_file_in_background(
    file_path: str,
    progressbar: ttk.Progressbar,
    progress_var: tk.DoubleVar,
    output_text: tk.Text,
    start_button: tk.Button,
) -> None:
    def worker() -> None:
        try:
            results: Dict[str, int] = process_excel_file(
                file_path,
                progress_callback=lambda v: progressbar.after(
                    0, update_progress, progressbar, progress_var, v
                ),
            )
            def write_results() -> None:
                output_text.config(state="normal")
                output_text.delete("1.0", tk.END)
                if not results:
                    output_text.insert(tk.END, "No sheets found.\n")
                else:
                    for sheet, count in results.items():
                        output_text.insert(tk.END, f"{sheet}: {count} rows\n")
                output_text.config(state="disabled")
                start_button.config(state="normal")
            output_text.after(0, write_results)
        except Exception as e:
            def show_err() -> None:
                messagebox.showerror("Error", str(e))
                start_button.config(state="normal")
            progressbar.after(0, show_err)

    start_button.config(state="disabled")
    threading.Thread(target=worker, daemon=True).start()

def create_main_window() -> tk.Tk:
    root = tk.Tk()
    root.title("XLSX Reader")

    frm = ttk.Frame(root, padding=12)
    frm.grid(sticky="nsew")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    ttk.Label(frm, text="Excel file:").grid(row=0, column=0, sticky="w", padx=(0, 8))
    file_entry = ttk.Entry(frm, width=50)
    file_entry.grid(row=0, column=1, sticky="ew")
    frm.columnconfigure(1, weight=1)
    ttk.Button(frm, text="Browse", command=lambda: select_excel_file(file_entry)).grid(row=0, column=2, padx=(8, 0))

    progress_var = tk.DoubleVar(value=0.0)
    progress = ttk.Progressbar(frm, orient="horizontal", mode="determinate", variable=progress_var, maximum=1.0)
    progress.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(12, 6))

    output = tk.Text(frm, height=10, width=60, state="disabled")
    output.grid(row=2, column=0, columnspan=3, sticky="nsew")
    frm.rowconfigure(2, weight=1)

    def on_start() -> None:
        path = file_entry.get().strip()
        if not path:
            messagebox.showwarning("Missing file", "Please choose an Excel file first.")
            return
        process_file_in_background(path, progress, progress_var, output, start_btn)

    start_btn = ttk.Button(frm, text="Process", command=on_start)
    start_btn.grid(row=3, column=0, columnspan=3, pady=(10, 0))

    return root

def run_app() -> None:
    root = create_main_window()
    root.mainloop()
