from __future__ import annotations

from typing import Callable, Dict, List, Optional

import pandas as pd


def get_sheet_names(file_path: str) -> List[str]:
    """
    Return all sheet names in the Excel workbook at file_path.
    """
    xls = pd.ExcelFile(file_path)
    return list(xls.sheet_names)


def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    """
    Return the number of data rows in the given sheet.
    Header is not counted as a row; len(df) is the expected behavior for tests.
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return int(len(df))


def process_excel_file(
    file_path: str,
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> Dict[str, int]:
    """
    Process all sheets and return a mapping of {sheet_name: row_count}.

    progress_callback, if provided, is called as progress_callback(done, total).
    """
    names = get_sheet_names(file_path)
    total = len(names)
    results: Dict[str, int] = {}
    for i, name in enumerate(names, start=1):
        results[name] = get_sheet_row_count(file_path, name)
        if progress_callback:
            progress_callback(i, total)
    return results
