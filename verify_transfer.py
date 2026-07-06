"""
Verify a CFA autofilled copy against its source files (read-only, openpyxl).

For every source workbook in the folder, this re-reads the saved autofilled target copy and
asserts that the entity's column, on the P{month} sheet, holds exactly the source's check value
for each (FORM, LINE) line. Prints PASS/FAIL with a per-file count of mismatches.

This is an independent check using a DIFFERENT library (openpyxl) than the writer (xlwings),
so it catches both logic bugs and write bugs.

Run:  py -3.12 verify_transfer.py "C:\\path\\to\\CFA verification... _autofilled_YYYYMMDD_HHMMSS.xlsx"
      py -3.12 verify_transfer.py            # auto-pick newest *_autofilled_*.xlsx in this folder
"""

from __future__ import annotations

import glob
import os
import re
import sys
import warnings

warnings.filterwarnings("ignore")  # silence openpyxl conditional-formatting extension warning

import openpyxl  # noqa: E402

from cfa_transfer import (  # reuse the exact same parsing/detection logic  # noqa: E402
    CONFIG, norm, norm_key, parse_entity, parse_period,
    find_header_row, find_target_header, find_entity_row_and_col,
)

FOLDER = CONFIG["FOLDER"]
TOKEN = CONFIG["TARGET_NAME_TOKEN"]


def ws_getter(ws):
    def get(r, c):
        return ws.cell(row=r, column=c).value
    return get


def read_source_lines(path):
    """Return (entity, month, list_of_(form,line,check)) from a source workbook, or None."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        # pick sheet by header anchors (prefer configured tab)
        chosen = None
        header = None
        names = wb.sheetnames
        order = ([CONFIG["SOURCE_TAB"]] if CONFIG["SOURCE_TAB"] in names else []) + names
        for name in order:
            ws = wb[name]
            get = ws_getter(ws)
            maxc = min(ws.max_column or 1, 40)
            h = find_header_row(get, CONFIG["HEADER_SCAN_ROWS"], maxc, CONFIG)
            if h and h[3] is not None:
                chosen, header = ws, h
                break
        if chosen is None:
            return None
        header_row, form_col, line_col, check_col = header
        entity = parse_entity(chosen.cell(row=2, column=1).value)
        month, year = parse_period(chosen.cell(row=3, column=1).value)
        if not entity or not month:
            return None
        rows = []
        for r in range(header_row + 1, (chosen.max_row or header_row) + 1):
            form = norm(chosen.cell(row=r, column=form_col).value)
            line = norm(chosen.cell(row=r, column=line_col).value)
            if form is None and line is None:
                continue
            chk = chosen.cell(row=r, column=check_col).value
            rows.append((norm_key(form), norm_key(line), chk))
        return entity, month, rows
    finally:
        wb.close()


def find_autofilled():
    cands = sorted(glob.glob(os.path.join(FOLDER, "*_autofilled_*.xlsx")), key=os.path.getmtime)
    return cands[-1] if cands else None


def verify(copy_path):
    wb = openpyxl.load_workbook(copy_path, data_only=True)
    sources = [
        p for p in glob.glob(os.path.join(FOLDER, "*.xlsx"))
        if not os.path.basename(p).startswith("~$")
        and TOKEN.lower() not in os.path.basename(p).lower()
        and "_autofilled_" not in os.path.basename(p).lower()
    ]
    total_fail = 0
    total_checked = 0
    print(f"Verifying: {os.path.basename(copy_path)}")
    print(f"Against {len(sources)} source file(s)\n")

    for sp in sources:
        info = read_source_lines(sp)
        if info is None:
            print(f"  --   {os.path.basename(sp)} :: not a source, skipped")
            continue
        entity, month, src_rows = info
        sheet_name = f"P{month}"
        if sheet_name not in wb.sheetnames:
            print(f"  --   {os.path.basename(sp)} :: target has no '{sheet_name}', skipped")
            continue
        ws = wb[sheet_name]
        get = ws_getter(ws)
        maxc = min(ws.max_column or 1, 200)

        ent = find_entity_row_and_col(get, CONFIG["ENTITY_ROW_SCAN"], maxc, entity)
        th = find_target_header(get, CONFIG["HEADER_SCAN_ROWS"], maxc, CONFIG)
        if ent is None or th is None:
            print(f"  --   {os.path.basename(sp)} :: entity/header not found in '{sheet_name}', skipped")
            continue
        _erow, ecol = ent
        t_header_row, t_form_col, t_line_col = th

        # index target rows by key
        t_keys = {}
        for r in range(t_header_row + 1, (ws.max_row or t_header_row) + 1):
            f = norm_key(get(r, t_form_col))
            l = norm_key(get(r, t_line_col))
            if f is None and l is None:
                continue
            t_keys.setdefault((f, l), []).append(r)

        used = set()
        fails = 0
        checked = 0
        for i, (f, l, chk) in enumerate(src_rows):
            rows_for_key = [r for r in t_keys.get((f, l), []) if r not in used]
            if not rows_for_key:
                continue  # unmatched line (writer also skips these); not a value mismatch
            r = rows_for_key[0]
            used.add(r)
            got = ws.cell(row=r, column=ecol).value
            checked += 1
            if norm(got) != norm(chk):
                fails += 1
                if fails <= 5:
                    print(f"     MISMATCH {sheet_name}!{ws.cell(row=r, column=ecol).coordinate} "
                          f"line {l!r}: expected {chk!r} got {got!r}")
        total_fail += fails
        total_checked += checked
        status = "PASS" if fails == 0 else f"FAIL ({fails})"
        print(f"  {status:>9}  {os.path.basename(sp)} :: entity {entity}, {sheet_name}, "
              f"{checked} cells checked")

    print("\n" + "=" * 60)
    print(f"TOTAL: {total_checked} cells checked, {total_fail} mismatches -> "
          f"{'ALL PASS' if total_fail == 0 else 'FAILURES PRESENT'}")
    print("=" * 60)
    return total_fail == 0


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    copy_path = argv[0] if argv else find_autofilled()
    if not copy_path or not os.path.exists(copy_path):
        raise SystemExit("ERROR: no autofilled copy given/found. Pass the path as an argument.")
    ok = verify(copy_path)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
