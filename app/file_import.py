"""
Utility to parse lead import files (CSV and Excel) into a list of dicts.

Supports:
- CSV (.csv) — any delimiter, BOM-safe
- Excel (.xlsx, .xls, .xlsm, .ods) — first sheet, all columns
- Tab-separated (.tsv)

Returns:
  headers  : list[str]  — original column names from the file (order preserved)
  rows     : list[dict] — one dict per data row, keyed by header name

All values are stripped strings; empty-string values are kept as "".
"""

from __future__ import annotations

import csv
import io
from typing import Any


def parse_import_file(
    contents: bytes,
    filename: str,
) -> tuple[list[str], list[dict[str, str]]]:
    """Parse bytes from an uploaded CSV or Excel file.

    Returns (headers, rows) where each row is a plain dict[str, str].
    Raises ValueError with a human-readable message on parse failure.
    """
    fname = (filename or "").lower()

    # ── Excel / ODS ──────────────────────────────────────────────
    if any(fname.endswith(ext) for ext in (".xlsx", ".xls", ".xlsm", ".ods")):
        return _parse_excel(contents)

    # ── CSV / TSV (and anything else) ────────────────────────────
    return _parse_csv(contents)


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _parse_csv(contents: bytes) -> tuple[list[str], list[dict[str, str]]]:
    """Parse CSV / TSV bytes into (headers, rows)."""
    # 1. Decode — try UTF-8 BOM first, then latin-1 as fallback
    try:
        text = contents.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = contents.decode("latin-1")

    if not text.strip():
        raise ValueError("The file is empty.")

    # 2. Sniff delimiter
    sniffer = csv.Sniffer()
    sample = text[:4096]
    try:
        dialect = sniffer.sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel  # default comma

    reader = csv.DictReader(io.StringIO(text), dialect=dialect)

    if not reader.fieldnames:
        raise ValueError("No column headers found in the file.")

    headers = [str(h).strip() for h in reader.fieldnames if str(h).strip()]
    rows: list[dict[str, str]] = []

    for raw_row in reader:
        row: dict[str, str] = {}
        for key, val in raw_row.items():
            k = str(key).strip() if key else ""
            if not k:
                continue
            row[k] = str(val).strip() if val is not None else ""
        rows.append(row)

    return headers, rows


def _parse_excel(contents: bytes) -> tuple[list[str], list[dict[str, str]]]:
    """Parse Excel / ODS bytes into (headers, rows) using openpyxl."""
    try:
        import openpyxl  # type: ignore
    except ImportError:
        raise ValueError(
            "Excel support requires openpyxl. Run: pip install openpyxl"
        )

    try:
        wb = openpyxl.load_workbook(
            io.BytesIO(contents), read_only=True, data_only=True
        )
    except Exception as exc:
        raise ValueError(f"Could not open Excel file: {exc}") from exc

    ws = wb.active  # first / active sheet

    rows_iter = ws.iter_rows(values_only=True)

    # First non-empty row = headers
    headers: list[str] = []
    for raw_row in rows_iter:
        candidates = [str(c).strip() if c is not None else "" for c in raw_row]
        if any(candidates):
            headers = candidates
            break

    if not headers:
        raise ValueError("The Excel file appears to be empty.")

    # Remove trailing empty headers
    while headers and not headers[-1]:
        headers.pop()

    if not headers:
        raise ValueError("No column headers found in the Excel file.")

    rows: list[dict[str, str]] = []
    for raw_row in rows_iter:
        cells = [str(c).strip() if c is not None else "" for c in raw_row]
        # Pad / trim to match header length
        cells = cells[: len(headers)]
        while len(cells) < len(headers):
            cells.append("")
        # Skip fully empty rows
        if not any(cells):
            continue
        rows.append(dict(zip(headers, cells)))

    wb.close()
    return headers, rows
