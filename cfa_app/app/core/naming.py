"""Parse CFA source filenames and identify folders to ignore.

The reported/export workbooks are named like 'CFA 2605 Reported 003 USD.xlsx' (also tolerant of
'CFA 2505 059.xlsx'). The 'CFA <YYPP>' token is the reliable year+period source, independent of the
(wildly inconsistent) folder names. Macro templates named 'CFA - 2025_P..' , 'CFA light', 'TFA',
Flash/Dashboard/Financials/Notes and non-Excel files are NOT sources.
"""

from __future__ import annotations

import re

# Folders whose contents are superseded / not to be scanned.
IGNORE_DIRS = {"old", "archive", "_archive", "olds", "bak"}

CURRENCIES = {"USD", "EUR", "EGP", "MYR", "GBP", "CAD", "CHF", "JPY", "CNY", "INR",
              "AUD", "SGD", "HKD", "MXN", "BRL", "ZAR", "SEK", "NOK", "DKK", "PLN", "TRY", "AED"}

# 'CFA' + 4-digit YYPP token (no dash => excludes 'CFA - 2025_P..').
_CANDIDATE_RE = re.compile(r"^CFA\s+(\d{2})(\d{2})\b", re.IGNORECASE)
_ENTITY_RE = re.compile(r"\b(\d{3}[A-Za-z]?)\b")
_YEAR_RE = re.compile(r"\b(20\d{2})\b")


def is_ignored_dir(name: str) -> bool:
    return name.strip().lower() in IGNORE_DIRS


def parse_cfa_name(filename: str):
    """Return (year, period, entity, currency) for a CFA source file, or None if not one.

    currency defaults to 'USD' when the name has no currency token.
    """
    stem = filename.rsplit(".", 1)[0]
    if "light" in stem.lower():
        return None
    m = _CANDIDATE_RE.match(stem)
    if not m:
        return None
    year = 2000 + int(m.group(1))
    period = int(m.group(2))
    rest = stem[m.end():]
    ent = _ENTITY_RE.search(rest)
    if not ent:
        return None
    entity = ent.group(1).upper()
    currency = "USD"
    for tok in re.findall(r"\b([A-Za-z]{3})\b", rest):
        if tok.upper() in CURRENCIES:
            currency = tok.upper()
            break
    return year, period, entity, currency


def parse_year_from_text(text: str):
    """Extract a 4-digit 20xx year from arbitrary text (e.g. the master filename)."""
    if not text:
        return None
    m = _YEAR_RE.search(text)
    return int(m.group(1)) if m else None


_FULL_YEAR_RE = re.compile(r"^(20\d{2})$")
_PPREFIX_RE = re.compile(r"^P\s*0?(\d{1,2})\b", re.IGNORECASE)
_LEADING_DIGITS_RE = re.compile(r"^(\d+)")


def folder_year(name: str):
    """If a folder name IS a bare year (e.g. '2026'), return it; else None."""
    m = _FULL_YEAR_RE.match(name.strip())
    return int(m.group(1)) if m else None


def folder_period(name: str):
    """Best-effort period number a folder name encodes, or None if it isn't a period folder.

    Handles 'P05'/'P5'/'P05 2025', '05 - May 2025'/'05_May'/'09 September', and 'YYPP ...'
    (e.g. '2505 May 2025' -> 5). A 3-digit lead is an ENTITY code (e.g. '003', '059 - …') -> None,
    and a 4-digit lead that isn't a valid period (e.g. a year '2025') -> None.
    """
    s = name.strip()
    m = _PPREFIX_RE.match(s)
    if m:
        p = int(m.group(1))
        return p if 1 <= p <= 12 else None
    m = _LEADING_DIGITS_RE.match(s)
    if m:
        d = m.group(1)
        if len(d) <= 2:
            p = int(d)
        elif len(d) == 4:          # YYPP -> period is the last two digits
            p = int(d[2:])
        else:                      # 3-digit entity code, or other -> not a period folder
            return None
        return p if 1 <= p <= 12 else None
    return None
