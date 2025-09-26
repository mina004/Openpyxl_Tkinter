# xlsx_reader/excel_processor.py
from __future__ import annotations

from typing import Callable, Dict, List, Optional

from openpyxl import load_workbook


def get_sheet_names(file_path: str) -> List[str]:
    """
    Return all sheet names in the Excel workbook at file_path.

    Uses openpyxl and closes the workbook explicitly.
    Raises FileNotFoundError if the file doesn't exist.
    """
    wb = load_workbook(filename=file_path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    """
    Return the number of data rows in the given sheet, EXCLUDING the header row.

    We treat the first row as a header.
    Count rows from row 2 onward that contain at least one non-empty value.
    """
    wb = load_workbook(filename=file_path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name]
        count = 0
        # start from row 2 (exclude header)
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                count += 1
        return count
    finally:
        wb.close()


def process_excel_file(
    file_path: str,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
) -> Dict[str, int]:
    """
    Process all sheets and return a mapping of {sheet_name: row_count}.

    If provided, progress_callback is called as:
        progress_callback(current_index, total_sheets, sheet_name)
    where current_index starts at 1.
    """
    sheet_names = get_sheet_names(file_path)
    total = len(sheet_names)
    results: Dict[str, int]
