"""Select the source file per entity+currency for a given year+period from a flat file list.

Input `files` is a list of dicts with at least: name, id, last_modified (ISO string). Files come
from a recursive SharePoint walk that already skips OLD/archive folders. Selection:
  - keep only files whose CFA <YYPP> token matches the requested year + period,
  - group by (entity, currency),
  - within a group choose by NAME (timestamps are unreliable after a bulk copy), highest wins:
      1. 'final' in the name,
      2. 'POST AUDIT' in the name,
      3. highest 'v{n}' (v5 > v4 …),
      4. .xlsx over .xlsm,
      5. highest trailing counter ('(2)', 'USD1'),
      6. latest last_modified.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .naming import CURRENCIES, parse_cfa_name

_V_RE = re.compile(r"\bv(\d+)\b")                 # v5, _v2  (word-boundary so 'revtab5' won't match)
_POST_AUDIT_RE = re.compile(r"post[\s_-]?audit")  # POST AUDIT / post-audit / postaudit
_PAREN_NUM_RE = re.compile(r"\((\d{1,3})\)\s*$")  # trailing "(2)"
_TAIL_NUM_RE = re.compile(r"[A-Za-z](\d{1,3})\s*$")  # digits glued after a letter, e.g. "USD1"
# 'CFA <YYPP> [Reported] <ENTITY>' prefix — anything after (past an optional currency) is a suffix.
_BASE_RE = re.compile(r"^CFA\s+\d{4}\s+(?:Reported\s+)?(\d{3}[A-Za-z]?)\b", re.IGNORECASE)


def _has_suffix(name: str) -> int:
    """1 if the filename has extra text after 'CFA <YYPP> [Reported] <ENTITY> [<CUR>]'.

    e.g. 'CFA 2606 Reported 181L USD rev.xlsx' -> 1, 'CFA 2606 Reported 181L USD.xlsx' -> 0.
    Any trailing text (rev / R1 / a date / after CN / …) counts, so it beats the bare name.
    """
    stem = name.rsplit(".", 1)[0].strip()
    m = _BASE_RE.match(stem)
    if not m:
        return 0
    parts = stem[m.end():].strip().split()
    if parts and parts[0].upper() in CURRENCIES:
        parts = parts[1:]                     # drop the currency token; the rest is the suffix
    return 1 if parts else 0


@dataclass
class Selection:
    entity: str
    currency: str
    item_id: str
    name: str
    last_modified: str
    folder_path: str = ""
    web_url: str = ""

    @property
    def full_path(self) -> str:
        return f"{self.folder_path.rstrip('/')}/{self.name}" if self.folder_path else self.name


def version_key(name: str, last_modified: str = ""):
    """Ranking key for choosing among versions of the same entity+currency (higher wins).

    Order: final > POST AUDIT > highest v{n} > has-any-suffix > .xlsx-over-.xlsm >
           highest trailing counter > mtime.
    'has-any-suffix' means the name carries extra text (rev / a date / after CN / …) — such a file
    is preferred over the bare 'CFA <YYPP> Reported <ENTITY> <CUR>.xlsx'.
    """
    base = name.rsplit(".", 1)[0]
    low = base.lower()

    final = 1 if "final" in low else 0
    post_audit = 1 if _POST_AUDIT_RE.search(low) else 0

    m = _V_RE.search(low)
    vnum = int(m.group(1)) if m else -1

    has_suffix = _has_suffix(name)                     # any trailing text beats the bare name
    ext = 1 if name.lower().endswith(".xlsx") else 0   # prefer .xlsx over .xlsm

    trailing = -1
    m = _PAREN_NUM_RE.search(base)
    if m:
        trailing = int(m.group(1))
    else:
        m = _TAIL_NUM_RE.search(base)
        if m:
            trailing = int(m.group(1))

    return (final, post_audit, vnum, has_suffix, ext, trailing, last_modified or "")


def select_sources(files: list[dict], year: int, period: int) -> list[Selection]:
    groups: dict[tuple, list] = {}
    for f in files:
        name = f.get("name", "")
        if not name.lower().endswith((".xlsx", ".xlsm")):
            continue
        parsed = parse_cfa_name(name)
        if parsed is None:
            continue
        fyear, fperiod, entity, currency = parsed
        if fyear != year or fperiod != period:
            continue
        groups.setdefault((entity, currency), []).append(f)

    out = []
    for (entity, currency), fl in sorted(groups.items()):
        best = max(fl, key=lambda f: version_key(f.get("name", ""), f.get("last_modified") or ""))
        out.append(Selection(entity=entity, currency=currency, item_id=best.get("id"),
                             name=best.get("name", ""), last_modified=best.get("last_modified") or "",
                             folder_path=best.get("folder_path") or "",
                             web_url=best.get("webUrl") or ""))
    return out
