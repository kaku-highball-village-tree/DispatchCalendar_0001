#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Dispatch calendar Excel-to-TSV converter (CMD version)."""

from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def write_error_text(psz_excel_file_path: str, psz_error_message: str) -> str:
    """Write an error text file beside the source Excel file."""
    obj_excel_path: Path = Path(psz_excel_file_path)
    psz_error_file_name: str = f"{obj_excel_path.stem}_error.txt"
    obj_error_file_path: Path = obj_excel_path.with_name(psz_error_file_name)

    with obj_error_file_path.open(mode="w", encoding="utf-8", newline="\r\n") as obj_error_file:
        obj_error_file.write(psz_error_message)

    return str(obj_error_file_path)


def convert_excel_to_tsv(psz_excel_file_path: str) -> str:
    """Convert active sheet of an Excel file to UTF-8 TSV with CRLF line endings."""
    obj_excel_path: Path = Path(psz_excel_file_path)
    psz_tsv_file_name: str = f"{obj_excel_path.stem}.tsv"
    obj_tsv_file_path: Path = obj_excel_path.with_name(psz_tsv_file_name)

    obj_workbook: Any = load_workbook(filename=str(obj_excel_path), data_only=True)
    obj_active_sheet: Any = obj_workbook.active

    with obj_tsv_file_path.open(mode="w", encoding="utf-8", newline="\r\n") as obj_tsv_file:
        for obj_row in obj_active_sheet.iter_rows(values_only=True):
            list_row_values: list[str] = []
            for obj_cell_value in obj_row:
                if obj_cell_value is None:
                    psz_cell_text: str = ""
                else:
                    psz_cell_text = str(obj_cell_value)
                list_row_values.append(psz_cell_text)
            psz_tsv_line: str = "\t".join(list_row_values)
            obj_tsv_file.write(psz_tsv_line + "\n")

    obj_workbook.close()
    return str(obj_tsv_file_path)


def main() -> int:
    list_arguments: list[str] = sys.argv
    if len(list_arguments) != 2:
        psz_usage_message: str = "Usage: python DispatchCalendar_Cmd.py <excel_file_path>"
        print(psz_usage_message)
        return 1

    psz_excel_file_path: str = list_arguments[1]
    obj_excel_path: Path = Path(psz_excel_file_path)

    if not obj_excel_path.exists() or not obj_excel_path.is_file():
        psz_error_message: str = f"Input file not found: {psz_excel_file_path}"
        print(psz_error_message)
        write_error_text(psz_excel_file_path, psz_error_message)
        return 1

    if obj_excel_path.suffix.lower() != ".xlsx":
        psz_error_message = f"Only .xlsx files are supported: {psz_excel_file_path}"
        print(psz_error_message)
        write_error_text(psz_excel_file_path, psz_error_message)
        return 1

    try:
        psz_created_tsv_path: str = convert_excel_to_tsv(psz_excel_file_path)
        print(f"TSV created: {psz_created_tsv_path}")
        return 0
    except Exception as obj_exception:  # noqa: BLE001
        psz_traceback_text: str = traceback.format_exc()
        psz_error_message = (
            "Failed to convert Excel to TSV.\n"
            f"Input: {psz_excel_file_path}\n"
            f"Error: {obj_exception}\n\n"
            f"Traceback:\n{psz_traceback_text}"
        )
        print(psz_error_message)
        write_error_text(psz_excel_file_path, psz_error_message)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
