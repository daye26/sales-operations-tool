"""Shared, source-neutral helpers for tabular Excel and CSV values."""

import re
import unicodedata


def is_missing(value):
    return value is None or str(value).strip() == ""


def format_value(value):
    if is_missing(value):
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").strip()


def normalize_header(value):
    text = unicodedata.normalize("NFKC", format_value(value)).lower()
    text = text.replace("-", " ").replace("_", " ")
    return re.sub(r"\s+", " ", text).strip()


def text_key(value, remove_accents=True):
    text = format_value(value).upper()
    if remove_accents:
        text = unicodedata.normalize("NFD", text)
        text = "".join(character for character in text if unicodedata.category(character) != "Mn")
    return re.sub(r"\s+", " ", text).strip()


def code_key(value):
    return format_value(value).upper()


def vin_key(value):
    return code_key(value)


def row_value(row, indexes, field):
    index = indexes.get(field)
    return row[index] if index is not None and index < len(row) else None


def header_index(headers, header_aliases, field, required=True, error_prefix="Missing column"):
    aliases = {normalize_header(alias) for alias in header_aliases[field]}
    for index, header in enumerate(headers):
        if normalize_header(header) in aliases:
            return index
    if required:
        raise ValueError(f"{error_prefix} {field}. Headers: {headers}")
    return None


def build_indexes(headers, header_aliases, required_fields, optional_fields=None, error_prefix="Missing column"):
    indexes = {
        field: header_index(headers, header_aliases, field, error_prefix=error_prefix)
        for field in required_fields
    }
    for field in optional_fields or ():
        indexes[field] = header_index(
            headers,
            header_aliases,
            field,
            required=False,
            error_prefix=error_prefix,
        )
    return indexes


def max_required_col(indexes):
    values = [index for index in indexes.values() if index is not None]
    return max(values) + 1 if values else 0
