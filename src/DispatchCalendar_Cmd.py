#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Dispatch calendar Excel-to-TSV converter (CMD version)."""

from __future__ import annotations

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


def split_cell_text(psz_cell_text: str) -> list[str]:
    """Split a single cell by line breaks; return at least one element."""
    list_split_lines: list[str] = psz_cell_text.splitlines()
    if len(list_split_lines) == 0:
        return [""]
    return list_split_lines


def normalize_line_breaks_and_trim(psz_text: str) -> str:
    """Normalize line breaks and trim spaces for strict matching."""
    psz_normalized_text: str = psz_text.replace("\r\n", "\n").strip()
    return psz_normalized_text


def should_skip_row_by_first_column(list_source_row: list[str], list_skip_keywords: list[str]) -> bool:
    """Return True when first-column text matches any normalized skip keyword exactly."""
    if len(list_source_row) == 0:
        return False

    psz_first_column_text: str = normalize_line_breaks_and_trim(list_source_row[0])
    if psz_first_column_text == "":
        return False

    for psz_skip_keyword in list_skip_keywords:
        psz_normalized_skip_keyword: str = normalize_line_breaks_and_trim(psz_skip_keyword)
        if psz_first_column_text == psz_normalized_skip_keyword:
            return True

    return False


def normalize_rows_with_multiline_expansion(list_source_rows: list[list[str]]) -> list[list[str]]:
    """Expand multiline cells by physical row without carry-forward fill."""
    list_normalized_rows: list[list[str]] = []

    for list_source_row in list_source_rows:
        list_row_split_values: list[list[str]] = []
        iMaxLogicalRowCount: int = 1

        for psz_cell_text in list_source_row:
            list_split_values: list[str] = split_cell_text(psz_cell_text)
            list_row_split_values.append(list_split_values)
            if len(list_split_values) > iMaxLogicalRowCount:
                iMaxLogicalRowCount = len(list_split_values)

        for iLogicalRowIndex in range(iMaxLogicalRowCount):
            list_expanded_row: list[str] = []
            for iColumnIndex, list_split_values in enumerate(list_row_split_values):
                psz_raw_value: str = ""
                if iLogicalRowIndex < len(list_split_values):
                    psz_raw_value = list_split_values[iLogicalRowIndex]

                list_expanded_row.append(psz_raw_value)

            list_normalized_rows.append(list_expanded_row)

    return list_normalized_rows


def convert_excel_to_tsv(psz_excel_file_path: str) -> str:
    """Convert active sheet of an Excel file to UTF-8 TSV with CRLF line endings."""
    obj_excel_path: Path = Path(psz_excel_file_path)
    psz_tsv_file_name: str = f"{obj_excel_path.stem}.tsv"
    obj_tsv_file_path: Path = obj_excel_path.with_name(psz_tsv_file_name)

    obj_workbook: Any = load_workbook(filename=str(obj_excel_path), data_only=True)
    obj_active_sheet: Any = obj_workbook.active

    list_skip_keywords: list[str] = ["始業前点検"]

    list_source_rows: list[list[str]] = []
    for obj_row in obj_active_sheet.iter_rows(values_only=True):
        list_row_values: list[str] = []
        for obj_cell_value in obj_row:
            if obj_cell_value is None:
                psz_cell_text: str = ""
            else:
                psz_cell_text = str(obj_cell_value)
            list_row_values.append(psz_cell_text)

        if should_skip_row_by_first_column(list_row_values, list_skip_keywords):
            continue

        list_source_rows.append(list_row_values)

    list_normalized_rows: list[list[str]] = normalize_rows_with_multiline_expansion(list_source_rows)

    with obj_tsv_file_path.open(mode="w", encoding="utf-8", newline="\r\n") as obj_tsv_file:
        for list_normalized_row in list_normalized_rows:
            psz_tsv_line: str = "\t".join(list_normalized_row)
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
