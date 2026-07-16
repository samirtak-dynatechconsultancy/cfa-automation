"""(FORM, LINE) rows to hide in every period sheet of the verification output.

Matching is by normalised key (norm_key), so minor whitespace/case differences don't matter. All
rows in a sheet whose (FORM, LINE) matches an entry here are hidden (Excel row hidden), on every
period sheet.
"""

from __future__ import annotations

from .detection import norm_key

_HIDE: list[tuple[str, str]] = [
    ("CFA 0.12", "1700 - Salaries/Wages"),
    ("CFA 1.31", "9261 - Short term receivables associated companies"),
    ("CFA 1.31", "9262 - Short term payables associated companies"),
    ("CFA 1.31", "9263 - Long term receivables associated companies"),
    ("CFA 1.31", "9264 - Long term payables associated companies"),
    ("CFA 1.32", "5130F99 - Associated companies"),
    ("CFA 1.32", "Movements Associated companies completed?"),
    ("CFA 1.32", "Impairment associated companies"),
    ("CFA 1.40", "5398I17 - Allowance obsolete and slow moving"),
    ("CFA 1.50", "5550D29 - Gross receivable associated companies"),
    ("CFA 2.40", "9220L52 - ST Redemptions on long term debt"),
    ("CFA 2.40", "9220L53 - Long term redemptions on long term debt"),
    ("CFA 2.60", "Limitation for the utilisation of existing carry forward losses?"),
    ("CFA 3.00", "Fair value land & buildings"),
    ("CFA 3.00", "Compliance with accounting manual"),
    ("CFA 3.01", "9255 - Operational lease <6 year"),
    ("CFA 3.10", "9181 - Market values forwards"),
    ("CFA 3.10", "9192 - Market values options"),
    ("CFA 3.10", "Embedded derivatives transaction risk?"),
    ("Rev by Country", "Gross third party revenues by country"),
    ("Add Rev Info", "Additional revenue information"),
    ("Audit fees", "Total audit fees"),
]

# Normalised (FORM, LINE) keys to hide.
HIDE_KEYS: set[tuple] = {(norm_key(f), norm_key(l)) for f, l in _HIDE}
