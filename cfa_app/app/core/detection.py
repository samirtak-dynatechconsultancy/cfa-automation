"""String normalisation and header/entity detection.

These helpers are engine-agnostic: they operate on a `get_cell(row, col)` callable, so the same
code works whether cell values come from openpyxl (this app) or xlwings (the original script).
Lifted, unchanged in behaviour, from the reference implementation.
"""

from __future__ import annotations

import re

from .models import DetectionConfig


def norm(v) -> str | None:
    """Normalise a cell value to a stripped string (or None)."""
    if v is None:
        return None
    s = str(v).strip()
    return s if s != "" else None


def norm_key(v) -> str | None:
    """Key form: stripped + collapsed internal whitespace, for label comparison."""
    s = norm(v)
    if s is None:
        return None
    return re.sub(r"\s+", " ", s)


def _contains(haystack, needle) -> bool:
    h = norm(haystack)
    return h is not None and needle.strip().upper() in h.upper()


def parse_entity(a2) -> str | None:
    """A2 = 'NUMBER: 055' -> '055'. Tolerant of spacing/case."""
    s = norm(a2)
    if not s:
        return None
    m = re.search(r"NUMBER\s*:?\s*([0-9A-Za-z]+)", s, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return None


def parse_period(a3):
    """A3 = 'PERIOD: 5 / 2026' -> (5, 2026). Returns (month, year) or (None, None)."""
    s = norm(a3)
    if not s:
        return None, None
    m = re.search(r"PERIOD\s*:?\s*(\d{1,2})\s*/\s*(\d{4})", s, re.IGNORECASE)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def find_header_row(get_cell, scan_rows: int, max_col: int, cfg: DetectionConfig):
    """Locate the table header row by finding FORM, LINE and CHECK header anchors.

    Returns (header_row, form_col, line_col, check_col) all 1-based, or None.
    check_col may be None (fine for the target, required for a source).
    """
    for r in range(1, scan_rows + 1):
        form_col = line_col = check_col = None
        for c in range(1, max_col + 1):
            v = get_cell(r, c)
            if v is None:
                continue
            if form_col is None and _contains(v, cfg.form_header):
                form_col = c
            elif line_col is None and _contains(v, cfg.line_header):
                line_col = c
            elif check_col is None and _contains(v, cfg.check_header):
                check_col = c
        if form_col and line_col:
            return r, form_col, line_col, check_col
    return None


def find_target_header(get_cell, scan_rows: int, max_col: int, cfg: DetectionConfig):
    """Target only needs FORM + LINE. Returns (header_row, form_col, line_col) or None."""
    res = find_header_row(get_cell, scan_rows, max_col, cfg)
    if res is None:
        return None
    header_row, form_col, line_col, _ = res
    return header_row, form_col, line_col


def find_entity_row_and_col(get_cell, scan_rows: int, max_col: int, entity):
    """Find the (row, col) in the top `scan_rows` whose stripped value equals the entity number.

    Returns (entity_row, entity_col) 1-based, or None.
    """
    target = str(entity).strip()
    for r in range(1, scan_rows + 1):
        for c in range(1, max_col + 1):
            v = get_cell(r, c)
            if v is None:
                continue
            if str(v).strip() == target:
                return r, c
    return None
