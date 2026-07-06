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

from .naming import parse_cfa_name

_V_RE = re.compile(r"\bv(\d+)\b")                 # v5, _v2  (word-boundary so 'revtab5' won't match)
_POST_AUDIT_RE = re.compile(r"post[\s_-]?audit")  # POST AUDIT / post-audit / postaudit
_PAREN_NUM_RE = re.compile(r"\((\d{1,3})\)\s*$")  # trailing "(2)"
_TAIL_NUM_RE = re.compile(r"[A-Za-z](\d{1,3})\s*$")  # digits glued after a letter, e.g. "USD1"


@dataclass
class Selection:
    entity: str
    currency: str
    item_id: str
    name: str
    last_modified: str
    folder_path: str = ""

    @property
    def full_path(self) -> str:
        return f"{self.folder_path.rstrip('/')}/{self.name}" if self.folder_path else self.name


def version_key(name: str, last_modified: str = ""):
    """Ranking key for choosing among versions of the same entity+currency (higher wins).

    Order: final > POST AUDIT > highest v{n} > .xlsx-over-.xlsm > highest trailing counter > mtime.
    """
    base = name.rsplit(".", 1)[0]
    low = base.lower()

    final = 1 if "final" in low else 0
    post_audit = 1 if _POST_AUDIT_RE.search(low) else 0

    m = _V_RE.search(low)
    vnum = int(m.group(1)) if m else -1

    ext = 1 if name.lower().endswith(".xlsx") else 0   # prefer .xlsx over .xlsm

    trailing = -1
    m = _PAREN_NUM_RE.search(base)
    if m:
        trailing = int(m.group(1))
    else:
        m = _TAIL_NUM_RE.search(base)
        if m:
            trailing = int(m.group(1))

    return (final, post_audit, vnum, ext, trailing, last_modified or "")


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
                             folder_path=best.get("folder_path") or ""))
    return out
