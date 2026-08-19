"""Helpers for memory-efficient, fresh Excel output workbooks."""

from datetime import date
import os
from pathlib import Path
import tempfile

from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter


def calculate_column_widths(headers, rows, row_values, format_value):
    widths = [len(format_value(header)) for header in headers]
    for row in rows:
        for column_index, value in enumerate(row_values(row), start=1):
            if column_index > len(widths):
                widths.append(0)
            widths[column_index - 1] = max(widths[column_index - 1], len(format_value(value)))
    return widths


def prepare_worksheet(worksheet, headers, widths, maximum_width):
    worksheet.freeze_panes = "A2"
    for column_index, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(width + 2, 10), maximum_width
        )
    worksheet.append(headers)


def append_row(worksheet, values, short_date_format):
    cells = []
    for value in values:
        if isinstance(value, date):
            cell = WriteOnlyCell(worksheet, value=value)
            cell.number_format = short_date_format
            cells.append(cell)
        else:
            cells.append(value)
    worksheet.append(cells)


def save_workbook_atomically(workbook, output_path):
    """Replace an output only after its complete workbook was written successfully."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            dir=output_path.parent,
            prefix=f".{output_path.stem}.",
            suffix=output_path.suffix or ".xlsx",
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
        workbook.save(temporary_path)
        os.replace(temporary_path, output_path)
        temporary_path = None
    finally:
        if temporary_path is not None:
            try:
                temporary_path.unlink()
            except FileNotFoundError:
                pass
