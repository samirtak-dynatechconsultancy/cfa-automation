"""Parse CFA source filenames and identify folders to ignore.

Only the exact reporting convention is treated as a source:
'CFA <YYPP> Reported <entity> <CUR>.xls[x|m]' (e.g. 'CFA 2606 Reported 045 USD.xlsx'). The
'CFA <YYPP>' token gives year+period, '<entity>' the company number, '<CUR>' the currency.
Anything else — a missing 'Reported'/currency token, or ANY extra text after the currency
('rev', 'final', 'v2', '(2)', a date), plus macro templates ('CFA - 2025_P..', 'CFA light'),
Flash/Dashboard/Notes and non-Excel files — is NOT a source.
"""

from __future__ import annotations

import re

# Folders whose contents are superseded / not to be scanned.
IGNORE_DIRS = {"old", "archive", "_archive", "olds", "bak"}

# Region folders sit directly under the year (Year > Region > Entity > P{no.} > file). These are the
# region groups the master verification sheet is organised into (row above the entity numbers).
KNOWN_REGIONS = {"AMS", "APAC", "EMEA", "HOLDINGS"}
_REGION_CANON = {"HOLDING": "HOLDINGS"}

CURRENCIES = {"USD", "EUR", "EGP", "MYR", "GBP", "CAD", "CHF", "JPY", "CNY", "INR",
              "AUD", "SGD", "HKD", "MXN", "BRL", "ZAR", "SEK", "NOK", "DKK", "PLN", "TRY", "AED"}

# Exact canonical source name: 'CFA <YYPP> Reported <entity> <CUR>' anchored end-to-end, so any
# trailing suffix (rev / final / v2 / (2) / a date) or a missing token disqualifies the file.
# The entity must be DIGITS ONLY (e.g. 045, 181, 278); a lettered entity like 181L / 604P is skipped.
_NAME_RE = re.compile(r"^CFA\s+(\d{2})(\d{2})\s+Reported\s+(\d{3})\s+([A-Za-z]{3})$",
                      re.IGNORECASE)
_YEAR_RE = re.compile(r"\b(20\d{2})\b")


def is_ignored_dir(name: str) -> bool:
    return name.strip().lower() in IGNORE_DIRS


def parse_cfa_name(filename: str):
    """Return (year, period, entity, currency) for a canonically-named CFA source, else None.

    Accepts ONLY 'CFA <YYPP> Reported <entity> <CUR>.xls[x|m]' with nothing after the currency
    (e.g. 'CFA 2606 Reported 045 USD.xlsx'). The entity must be digits only — a lettered entity
    such as 181L or 604P is rejected. The currency must be a known one; any suffix, a missing
    'Reported', or a missing/unknown currency returns None.
    """
    stem = filename.rsplit(".", 1)[0].strip()
    m = _NAME_RE.match(stem)
    if not m:
        return None
    currency = m.group(4).upper()
    if currency not in CURRENCIES:
        return None
    return 2000 + int(m.group(1)), int(m.group(2)), m.group(3).upper(), currency


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


def canon_region(s: str) -> str:
    """Normalise a region label for matching ('Holdings' -> 'HOLDINGS', 'holding' -> 'HOLDINGS')."""
    u = re.sub(r"\s+", " ", str(s or "").strip()).upper()
    return _REGION_CANON.get(u, u)


def folder_region(folder_path: str, year: int | None = None):
    """Region = the folder directly inside the year folder. '/2026/EMEA/055/P05' -> 'EMEA'.

    Returns the raw region token as it appears in the path (or None). The caller matches it to the
    master's region header case-insensitively via canon_region().
    """
    if not folder_path:
        return None
    parts = [p for p in re.split(r"[\\/]+", str(folder_path)) if p.strip()]
    if year is not None:
        ystr = str(year)
        for i, p in enumerate(parts):
            if p.strip() == ystr or folder_year(p) == year:
                return parts[i + 1].strip() if i + 1 < len(parts) else None
    for p in parts:                                   # fallback: any part that is a known region
        if canon_region(p) in KNOWN_REGIONS:
            return p.strip()
    return None


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
