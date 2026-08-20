"""Worksheet selection helpers for workbooks whose active sheet can vary."""

import re
import unicodedata


def normalize_header(value):
    text = "" if value is None else str(value).strip()
    text = unicodedata.normalize("NFKC", text).lower()
    text = text.replace("-", " ").replace("_", " ")
    return re.sub(r"\s+", " ", text).strip()


def select_active_then_sheet1(workbook, header_aliases, required_columns):
    """Use the active worksheet when it has the expected headers, else Sheet1."""
    candidates = [workbook.active]
    if "Sheet1" in workbook.sheetnames and workbook["Sheet1"] is not workbook.active:
        candidates.append(workbook["Sheet1"])

    expected_headers = [
        {normalize_header(alias) for alias in header_aliases[column]}
        for column in required_columns
    ]
    for worksheet in candidates:
        reset_dimensions = getattr(worksheet, "reset_dimensions", None)
        if reset_dimensions is not None:
            reset_dimensions()
        try:
            headers = next(worksheet.iter_rows(max_row=1, values_only=True))
        except StopIteration:
            continue
        available_headers = {normalize_header(header) for header in headers}
        if all(aliases & available_headers for aliases in expected_headers):
            return worksheet

    reset_dimensions = getattr(workbook.active, "reset_dimensions", None)
    if reset_dimensions is not None:
        reset_dimensions()
    return workbook.active
