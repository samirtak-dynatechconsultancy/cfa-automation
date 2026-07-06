"""
CFA Consistency-Check Transfer
==============================

Automates the manual copy job: for every entity source workbook in a folder, read the
consistency-check results (OK / NOT EQUAL / -----) from the consistency-check tab (col D,
rows ~10-210) and write them into the matching entity COLUMN of the shared
`CFA verification_Pn YYYY.xlsx` master, on the correct period sheet (P5, P4, ...).

Engine: xlwings (drives the installed desktop Excel) so ALL formatting -- including the red
conditional-formatting highlights on "NOT EQUAL" -- is preserved. openpyxl would strip them.

Design goals (per approved plan):
- Process a WHOLE FOLDER of entity files in one run.
- The target master lives IN THE SAME folder -> identify and exclude it from the source set.
- Read entity number + period from INSIDE each source file (A2/A3), never from the filename.
- Auto-detect the sheet / header row / columns via header anchors so a renamed tab, an
  inserted line, or a shifted column does not break the pipeline.
- Row-position transfer WITH per-row label verification (handles the 5 duplicate (FORM,LINE)
  keys and degrades gracefully on template drift).
- Save to a TIMESTAMPED COPY (original master untouched). Refresh the whole entity column.
- Missing entity / period -> skip + log, keep going. Print a done/skipped summary.

Run:  py -3.12 cfa_transfer.py            (uses CONFIG below)
      py -3.12 cfa_transfer.py --folder "C:\\path\\to\\folder"
      py -3.12 cfa_transfer.py --verify-only "C:\\path\\to\\autofilled_copy.xlsx"
"""

from __future__ import annotations

import argparse
import datetime as _dt
import glob
import os
import re
import sys
from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# CONFIG  (edit these; nothing below should need changing for a normal run)
# ---------------------------------------------------------------------------

CONFIG = {
    # Folder that holds BOTH the entity source files AND the target master.
    "FOLDER": os.path.dirname(os.path.abspath(__file__)),

    # Substring (case-insensitive) used to recognise the target master workbook.
    "TARGET_NAME_TOKEN": "verification",

    # Preferred consistency-check tab name in the SOURCE files. If it is renamed,
    # the script falls back to auto-detecting the sheet by its header row.
    "SOURCE_TAB": "G",

    # Header anchor strings used to auto-locate the table (compared case-insensitively,
    # whitespace-stripped, as substrings). Order independent.
    "FORM_HEADER": "FORM",
    "LINE_HEADER": "LINE",
    "CHECK_HEADER": "EQUAL TO",   # the "  EQUAL TO ?" column = the value to transfer

    # Search bounds (generous safety caps; actual extents are auto-detected within these).
    "HEADER_SCAN_ROWS": 20,       # scan first N rows for the header row
    "ENTITY_ROW_SCAN": 20,        # scan first N rows of target for the entity-number row
    "MAX_DATA_ROWS": 400,         # never look past this many rows below the header

    # Make the driven Excel instance visible (useful while debugging). False = headless-ish.
    "EXCEL_VISIBLE": False,
}

# Values that legitimately appear in the check column (informational only; we copy verbatim).
KNOWN_CHECK_VALUES = {"OK", "NOT EQUAL", "-----"}


# ---------------------------------------------------------------------------
# Small data holders
# ---------------------------------------------------------------------------

@dataclass
class SourceData:
    path: str
    entity: str
    month: int
    year: int
    sheet_name: str
    header_row: int           # 1-based
    form_col: int             # 1-based
    line_col: int             # 1-based
    check_col: int            # 1-based
    rows: list = field(default_factory=list)  # list of (form, line, check_value), data order


@dataclass
class FileResult:
    path: str
    entity: str | None = None
    period: str | None = None
    status: str = "skipped"       # "done" | "skipped"
    written: int = 0
    skipped_lines: int = 0
    messages: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Generic helpers (string normalisation, header detection)
# ---------------------------------------------------------------------------

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


def find_header_row(get_cell, scan_rows, max_col, cfg):
    """
    Locate the table header row by finding a row that contains the FORM, LINE and CHECK
    header anchors. Returns (header_row, form_col, line_col, check_col) all 1-based, or None.
    `get_cell(r, c)` returns the cell value.
    """
    for r in range(1, scan_rows + 1):
        form_col = line_col = check_col = None
        for c in range(1, max_col + 1):
            v = get_cell(r, c)
            if v is None:
                continue
            if form_col is None and _contains(v, cfg["FORM_HEADER"]):
                form_col = c
            elif line_col is None and _contains(v, cfg["LINE_HEADER"]):
                line_col = c
            elif check_col is None and _contains(v, cfg["CHECK_HEADER"]):
                check_col = c
        if form_col and line_col:
            # check_col is required for the source; for the target it is irrelevant (None ok).
            return r, form_col, line_col, check_col
    return None


def find_target_header(get_cell, scan_rows, max_col, cfg):
    """Target only needs FORM + LINE columns. Returns (header_row, form_col, line_col) or None."""
    res = find_header_row(get_cell, scan_rows, max_col, cfg)
    if res is None:
        return None
    header_row, form_col, line_col, _ = res
    return header_row, form_col, line_col


def find_entity_row_and_col(get_cell, scan_rows, max_col, entity):
    """
    Find the row that carries entity numbers and the column matching `entity` (stripped compare).
    Returns (entity_row, entity_col) 1-based, or None. The entity row is the one in the top
    `scan_rows` that contains a stripped-cell equal to the entity number.
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


# ---------------------------------------------------------------------------
# Source reading (xlwings)
# ---------------------------------------------------------------------------

def read_source(book, cfg) -> SourceData | None:
    """Read one opened source workbook. Returns SourceData, or None if it is not a valid source."""
    # Pick the consistency-check sheet: preferred name, else auto-detect by header.
    sheet = None
    names = [s.name for s in book.sheets]
    if cfg["SOURCE_TAB"] in names:
        sheet = book.sheets[cfg["SOURCE_TAB"]]

    def cell_getter(sht):
        # Cache the used block once for speed.
        last_row = min(sht.cells.last_cell.row, cfg["HEADER_SCAN_ROWS"])
        last_col = min(sht.cells.last_cell.column, 40)
        block = sht.range((1, 1), (max(last_row, 1), max(last_col, 1))).value
        if not isinstance(block, list):
            block = [[block]]
        elif block and not isinstance(block[0], list):
            block = [block]

        def get(r, c):
            if 1 <= r <= len(block) and 1 <= c <= len(block[r - 1]):
                return block[r - 1][c - 1]
            return None
        return get, last_col

    header = None
    if sheet is not None:
        get, last_col = cell_getter(sheet)
        header = find_header_row(get, cfg["HEADER_SCAN_ROWS"], last_col, cfg)

    if header is None:
        # Auto-detect: scan every sheet for one whose header row has FORM/LINE/CHECK anchors.
        for sht in book.sheets:
            get, last_col = cell_getter(sht)
            h = find_header_row(get, cfg["HEADER_SCAN_ROWS"], last_col, cfg)
            if h and h[3] is not None:   # require the CHECK column for a source
                sheet, header = sht, h
                break

    if sheet is None or header is None or header[3] is None:
        return None  # not a consistency-check workbook

    header_row, form_col, line_col, check_col = header

    # Metadata from A2 / A3 of the SAME sheet.
    a2 = sheet.range((2, 1)).value
    a3 = sheet.range((3, 1)).value
    entity = parse_entity(a2)
    month, year = parse_period(a3)
    if not entity or not month:
        return None  # missing identifying metadata -> not a usable source

    # Read the data block in one shot: FORM, LINE, CHECK columns, header_row+1 .. end.
    first = header_row + 1
    last = first + cfg["MAX_DATA_ROWS"]
    lo_col = min(form_col, line_col, check_col)
    hi_col = max(form_col, line_col, check_col)
    block = sheet.range((first, lo_col), (last, hi_col)).value
    if block and not isinstance(block[0], list):
        block = [block]

    rows = []
    last_nonempty = -1
    for i, raw in enumerate(block or []):
        form = raw[form_col - lo_col]
        line = raw[line_col - lo_col]
        chk = raw[check_col - lo_col]
        if norm(form) is None and norm(line) is None:
            rows.append((None, None, None))  # keep blanks as separators
        else:
            rows.append((norm(form), norm(line), chk))
            last_nonempty = i
    rows = rows[: last_nonempty + 1]  # trim trailing empties

    return SourceData(
        path=book.fullname, entity=entity, month=month, year=year,
        sheet_name=sheet.name, header_row=header_row, form_col=form_col,
        line_col=line_col, check_col=check_col, rows=rows,
    )


# ---------------------------------------------------------------------------
# Transfer into the target
# ---------------------------------------------------------------------------

def transfer_to_target(target_book, src: SourceData, cfg, result: FileResult) -> bool:
    """Write one source's check column into the matching target sheet+column. Returns True if done."""
    sheet_name = f"P{src.month}"
    names = [s.name for s in target_book.sheets]
    if sheet_name not in names:
        result.messages.append(f"period sheet '{sheet_name}' not found in target -> skipped")
        return False
    tsheet = target_book.sheets[sheet_name]

    # Read enough of the target to locate the entity row/col and the FORM/LINE block.
    last_row = min(tsheet.cells.last_cell.row, cfg["MAX_DATA_ROWS"] + cfg["HEADER_SCAN_ROWS"])
    last_col = min(tsheet.cells.last_cell.column, 200)
    block = tsheet.range((1, 1), (last_row, last_col)).value
    if block and not isinstance(block[0], list):
        block = [block]

    def tget(r, c):
        if 1 <= r <= len(block) and 1 <= c <= len(block[r - 1]):
            return block[r - 1][c - 1]
        return None

    # Entity column.
    ent = find_entity_row_and_col(tget, cfg["ENTITY_ROW_SCAN"], last_col, src.entity)
    if ent is None:
        result.messages.append(f"entity '{src.entity}' not found in '{sheet_name}' row scan -> skipped")
        return False
    _entity_row, entity_col = ent

    # Header row + FORM/LINE columns in the target.
    th = find_target_header(tget, cfg["HEADER_SCAN_ROWS"], last_col, cfg)
    if th is None:
        result.messages.append(f"could not locate FORM/LINE header in target '{sheet_name}' -> skipped")
        return False
    t_header_row, t_form_col, t_line_col = th

    # Build a (FORM,LINE) -> [rows] index for fallback matching, and a flat list for position.
    t_first = t_header_row + 1
    t_keys = {}            # key -> list of row numbers
    t_rowkey = {}          # row -> key
    for r in range(t_first, len(block) + 1):
        f = norm_key(tget(r, t_form_col))
        l = norm_key(tget(r, t_line_col))
        if f is None and l is None:
            continue
        key = (f, l)
        t_keys.setdefault(key, []).append(r)
        t_rowkey[r] = key

    # Transfer line by line. Source data row i (0-based) sits at source row src.header_row+1+i,
    # which corresponds, by template alignment, to target row t_first+i. Verify the label before
    # writing; on mismatch fall back to key search; on ambiguity/none -> skip that cell + log.
    writes = []  # (target_row, value)
    used_target_rows = set()
    for i, (form, line, chk) in enumerate(src.rows):
        if form is None and line is None:
            continue  # separator
        src_key = (norm_key(form), norm_key(line))
        candidate = t_first + i

        chosen = None
        if t_rowkey.get(candidate) == src_key and candidate not in used_target_rows:
            chosen = candidate
        else:
            # Fallback: key search, preferring an unused row, then nearest to candidate.
            rows_for_key = [r for r in t_keys.get(src_key, []) if r not in used_target_rows]
            if len(rows_for_key) == 1:
                chosen = rows_for_key[0]
            elif len(rows_for_key) > 1:
                chosen = min(rows_for_key, key=lambda r: abs(r - candidate))
                result.messages.append(
                    f"line {src_key[1]!r}: duplicate key, matched by nearest row {chosen}")
            else:
                result.messages.append(
                    f"line {src_key[1]!r} (form {src_key[0]!r}) not found in target -> cell skipped")
                result.skipped_lines += 1
                continue

        used_target_rows.add(chosen)
        writes.append((chosen, chk))

    if not writes:
        result.messages.append("no lines matched -> nothing written")
        return False

    # Write each value into (entity_col, target_row). Group contiguous rows for fewer COM calls.
    writes.sort()
    i = 0
    while i < len(writes):
        j = i
        while j + 1 < len(writes) and writes[j + 1][0] == writes[j][0] + 1:
            j += 1
        col_vals = [[writes[k][1]] for k in range(i, j + 1)]
        r0 = writes[i][0]
        r1 = writes[j][0]
        tsheet.range((r0, entity_col), (r1, entity_col)).value = col_vals
        i = j + 1

    result.written = len(writes)
    return True


# ---------------------------------------------------------------------------
# Folder discovery
# ---------------------------------------------------------------------------

def discover_files(folder, token):
    """Return (target_path, [source_paths]). Raises on 0 or >1 target candidates."""
    all_xlsx = [
        p for p in glob.glob(os.path.join(folder, "*.xlsx"))
        if not os.path.basename(p).startswith("~$")
    ]
    targets = [p for p in all_xlsx if token.lower() in os.path.basename(p).lower()]
    if len(targets) == 0:
        raise SystemExit(f"ERROR: no target master (name containing '{token}') found in {folder}")
    if len(targets) > 1:
        names = "\n  ".join(os.path.basename(t) for t in targets)
        raise SystemExit(
            f"ERROR: multiple files match the target token '{token}'; cannot pick the master:\n  {names}")
    target = targets[0]
    sources = [p for p in all_xlsx if p != target]
    return target, sources


# ---------------------------------------------------------------------------
# Main run
# ---------------------------------------------------------------------------

def run(folder, cfg):
    import xlwings as xw

    target_path, source_paths = discover_files(folder, cfg["TARGET_NAME_TOKEN"])
    print(f"Target master : {os.path.basename(target_path)}")
    print(f"Source files  : {len(source_paths)} found\n")

    results = []
    app = xw.App(visible=cfg["EXCEL_VISIBLE"], add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    save_path = None
    try:
        target_book = app.books.open(target_path, update_links=False, read_only=False)

        for sp in source_paths:
            res = FileResult(path=sp)
            sbook = None
            try:
                sbook = app.books.open(sp, update_links=False, read_only=True)
                src = read_source(sbook, cfg)
                if src is None:
                    res.messages.append("not a valid consistency-check source (no NUMBER/PERIOD/tab)")
                    results.append(res)
                    print(f"  SKIP  {os.path.basename(sp)} :: {res.messages[-1]}")
                    continue
                res.entity = src.entity
                res.period = f"P{src.month} ({src.month}/{src.year})"
                ok = transfer_to_target(target_book, src, cfg, res)
                res.status = "done" if ok else "skipped"
                tag = "DONE " if ok else "SKIP "
                extra = f"{res.written} lines" if ok else (res.messages[-1] if res.messages else "")
                print(f"  {tag} {os.path.basename(sp)} :: entity {src.entity}, P{src.month} :: {extra}")
                if res.skipped_lines:
                    print(f"        ({res.skipped_lines} individual lines skipped - see summary)")
            except Exception as e:  # noqa: BLE001 - one bad file must not kill the batch
                res.messages.append(f"error: {e!r}")
                print(f"  ERR   {os.path.basename(sp)} :: {e!r}")
            finally:
                if sbook is not None:
                    sbook.close()
            results.append(res)

        # Save a timestamped copy next to the target.
        ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(target_path)
        save_path = f"{base} _autofilled_{ts}{ext}"
        target_book.save(save_path)
        target_book.close()
    finally:
        app.quit()

    print_summary(results, save_path)
    return results, save_path


def print_summary(results, save_path):
    done = [r for r in results if r.status == "done"]
    skipped = [r for r in results if r.status != "done"]
    print("\n" + "=" * 70)
    print("RUN SUMMARY")
    print("=" * 70)
    print(f"Processed (done): {len(done)}   Skipped: {len(skipped)}")
    if save_path:
        print(f"Saved copy      : {save_path}")
    if done:
        print("\nDone:")
        for r in done:
            note = f"  ({r.skipped_lines} lines skipped)" if r.skipped_lines else ""
            print(f"  + {os.path.basename(r.path)}  entity {r.entity}  {r.period}  "
                  f"{r.written} lines{note}")
    if skipped:
        print("\nSkipped / issues:")
        for r in skipped:
            msg = r.messages[-1] if r.messages else "unknown"
            print(f"  - {os.path.basename(r.path)}  ({r.entity or '?'})  :: {msg}")
    # Surface any per-line warnings.
    warned = [r for r in results if r.messages and r.status == "done"]
    if warned:
        print("\nWarnings on processed files:")
        for r in warned:
            for m in r.messages:
                print(f"  ! {os.path.basename(r.path)} :: {m}")
    print("=" * 70)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv=None):
    p = argparse.ArgumentParser(description="Transfer CFA consistency checks into the verification master.")
    p.add_argument("--folder", default=CONFIG["FOLDER"], help="Folder with source files + target master.")
    p.add_argument("--token", default=CONFIG["TARGET_NAME_TOKEN"], help="Target filename token.")
    p.add_argument("--visible", action="store_true", help="Show Excel while running.")
    args = p.parse_args(argv)

    cfg = dict(CONFIG)
    cfg["FOLDER"] = args.folder
    cfg["TARGET_NAME_TOKEN"] = args.token
    if args.visible:
        cfg["EXCEL_VISIBLE"] = True

    run(cfg["FOLDER"], cfg)


if __name__ == "__main__":
    main()
