from __future__ import annotations
from typing import Callable, Dict, List, Optional
import pandas as pd

def get_sheet_names(file_path: str) -> List[str]:
    xls = pd.ExcelFile(file_path)
    return list(xls.sheet_names)

def get_sheet_row_count(file_path: str, sheet_name: str) -> int:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return int(len(df))

def process_excel_file(
    file_path: str,
    progress_callback: Optional[Callable[[float], None]] = None,
) -> Dict[str, int]:
    sheet_names = get_sheet_names(file_path)
    results: Dict[str, int] = {}
    total = len(sheet_names) or 1
    for i, name in enumerate(sheet_names, start=1):
        results[name] = get_sheet_row_count(file_path, name)
        if progress_callback:
            progress_callback(i / total)
    if progress_callback:
        progress_callback(1.0)
    return results
