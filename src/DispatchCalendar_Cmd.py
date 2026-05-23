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




def is_shifted_continuation_row(list_previous_row: list[str], list_current_row: list[str]) -> bool:
    """Detect left-shifted continuation rows that should map A/B into C/D."""
    if len(list_previous_row) < 4 or len(list_current_row) < 2:
        return False

    psz_previous_name: str = normalize_line_breaks_and_trim(list_previous_row[0])
    psz_previous_car_number: str = normalize_line_breaks_and_trim(list_previous_row[1])
    psz_previous_spare_car: str = normalize_line_breaks_and_trim(list_previous_row[2])
    psz_current_first_value: str = normalize_line_breaks_and_trim(list_current_row[0])
    psz_current_second_value: str = normalize_line_breaks_and_trim(list_current_row[1])

    if psz_previous_name == "" or psz_previous_car_number == "" or psz_previous_spare_car == "":
        return False

    if psz_current_first_value == "" or psz_current_second_value == "":
        return False

    for psz_remaining_value in list_current_row[2:]:
        if normalize_line_breaks_and_trim(psz_remaining_value) != "":
            return False

    return True


def shift_row_to_the_right(list_current_row: list[str], iShiftCount: int) -> list[str]:
    """Return a row shifted right by adding empty cells at the beginning."""
    return [""] * iShiftCount + list_current_row
def merge_continuation_rows(list_source_rows: list[list[str]]) -> list[list[str]]:
    """Merge continuation rows into previous row by joining values with newlines per column."""
    list_merged_rows: list[list[str]] = []

    for list_current_row in list_source_rows:
        b_has_previous_row: bool = len(list_merged_rows) > 0
        b_is_blank_first_column_continuation: bool = b_has_previous_row and normalize_line_breaks_and_trim(list_current_row[0]) == ""

        list_adjusted_current_row: list[str] = list_current_row[:]
        if b_has_previous_row and is_shifted_continuation_row(list_merged_rows[-1], list_current_row):
            list_adjusted_current_row = shift_row_to_the_right(list_current_row, 2)

        b_is_continuation_row: bool = b_is_blank_first_column_continuation or (list_adjusted_current_row is not list_current_row)

        if not b_is_continuation_row:
            list_merged_rows.append(list_current_row[:])
            continue

        list_previous_row: list[str] = list_merged_rows[-1]
        i_previous_column_count: int = len(list_previous_row)
        i_current_column_count: int = len(list_adjusted_current_row)
        i_max_column_count: int = i_previous_column_count if i_previous_column_count > i_current_column_count else i_current_column_count

        if i_previous_column_count < i_max_column_count:
            list_previous_row.extend([""] * (i_max_column_count - i_previous_column_count))

        for iColumnIndex in range(i_max_column_count):
            psz_current_value: str = ""
            if iColumnIndex < i_current_column_count:
                psz_current_value = list_adjusted_current_row[iColumnIndex]

            if normalize_line_breaks_and_trim(psz_current_value) == "":
                continue

            psz_previous_value: str = list_previous_row[iColumnIndex]
            if normalize_line_breaks_and_trim(psz_previous_value) == "":
                list_previous_row[iColumnIndex] = psz_current_value
            else:
                list_previous_row[iColumnIndex] = f"{psz_previous_value}\n{psz_current_value}"

    return list_merged_rows




def escape_newlines_in_cell_text(psz_cell_text: str) -> str:
    """Convert embedded newlines in cell text into literal \n sequences."""
    psz_escaped_text: str = psz_cell_text.replace("\r\n", "\n").replace("\n", "\\n")
    return psz_escaped_text


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
                psz_cell_text = escape_newlines_in_cell_text(str(obj_cell_value))
            list_row_values.append(psz_cell_text)

        if should_skip_row_by_first_column(list_row_values, list_skip_keywords):
            continue

        list_source_rows.append(list_row_values)

    list_normalized_rows: list[list[str]] = merge_continuation_rows(list_source_rows)

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
