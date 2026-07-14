"""openpyxl port of the CFA transfer + verify logic.

Algorithm is identical to the original xlwings script (position transfer with per-row label
verification, key-search fallback, nearest-row disambiguation for duplicate keys, graceful
per-line skip). Only the cell-access layer changed from COM/xlwings to openpyxl.

Sources are loaded with data_only=True so we read the CACHED computed value of the
"EQUAL TO ?" formula column (OK / NOT EQUAL / -----). The master is loaded normally so that
its conditional-formatting rules and other styling survive the round-trip; the red NOT-EQUAL
highlight is driven by a rule already present in the template and re-fires on the written value.
"""

from __future__ import annotations

import io
import re

import openpyxl
from openpyxl.styles import PatternFill

from .detection import (
    _contains,
    find_entity_row_and_col,
    find_header_row,
    find_target_header,
    norm,
    norm_key,
    parse_company,
    parse_entity,
    parse_period,
)
from .models import DetectionConfig, FileResult, SourceData, VerifyResult


# ---------------------------------------------------------------------------
# Source reading
# ---------------------------------------------------------------------------

def _ws_getter(ws):
    def get(r, c):
        return ws.cell(row=r, column=c).value
    return get


def read_source_from_bytes(data: bytes, filename: str, cfg: DetectionConfig) -> SourceData | None:
    """Parse one source workbook (bytes). Returns SourceData, or None if not a valid source."""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    try:
        names = wb.sheetnames
        order = ([cfg.source_tab] if cfg.source_tab in names else []) + names
        chosen = None
        header = None
        for name in order:
            ws = wb[name]
            get = _ws_getter(ws)
            maxc = min(ws.max_column or 1, 40)
            h = find_header_row(get, cfg.header_scan_rows, maxc, cfg)
            if h and h[3] is not None:      # require the CHECK column for a source
                chosen, header = ws, h
                break
        if chosen is None or header is None:
            return None

        header_row, form_col, line_col, check_col = header
        entity = parse_entity(chosen.cell(row=2, column=1).value)
        company = parse_company(chosen.cell(row=1, column=1).value)
        month, year = parse_period(chosen.cell(row=3, column=1).value)
        if not entity or not month:
            return None

        # Locate the '|difference|' column in the same header row (optional).
        diff_col = None
        for c in range(1, min(chosen.max_column or 1, 40) + 1):
            if _contains(chosen.cell(row=header_row, column=c).value, cfg.diff_header):
                diff_col = c
                break

        rows = []
        last_nonempty = -1
        first = header_row + 1
        last = min(first + cfg.max_data_rows, (chosen.max_row or first))
        for idx, r in enumerate(range(first, last + 1)):
            form = norm(chosen.cell(row=r, column=form_col).value)
            line = norm(chosen.cell(row=r, column=line_col).value)
            if form is None and line is None:
                rows.append((None, None, None, None))  # keep blanks as separators
            else:
                diff = chosen.cell(row=r, column=diff_col).value if diff_col else None
                rows.append((form, line, chosen.cell(row=r, column=check_col).value, diff))
                last_nonempty = idx
        rows = rows[: last_nonempty + 1]              # trim trailing empties

        return SourceData(
            filename=filename, entity=entity, month=month, year=year or 0,
            sheet_name=chosen.title, header_row=header_row, form_col=form_col,
            line_col=line_col, check_col=check_col, rows=rows, company=company or "",
        )
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Master loading / saving
# ---------------------------------------------------------------------------

_PSHEET_RE = re.compile(r"^P\d+$")
_ENTITY_CELL_RE = re.compile(r"^\d{3}[A-Za-z]?$")   # entity code like 055, 181L, 604P


def find_entity_header_row(get_cell, scan_rows: int, max_col: int):
    """Return the row that holds entity numbers (the one with the most entity-code cells)."""
    best_row, best_count = None, 0
    for r in range(1, scan_rows + 1):
        count = 0
        for c in range(1, max_col + 1):
            v = get_cell(r, c)
            if v is not None and _ENTITY_CELL_RE.match(str(v).strip()):
                count += 1
        if count > best_count:
            best_row, best_count = r, count
    return best_row if best_count >= 2 else None


def _next_entity_column(get_cell, entity_row: int, scan_cols: int) -> int:
    """First free column after the last entity code on the entity row."""
    last = 0
    for c in range(1, scan_cols + 1):
        v = get_cell(entity_row, c)
        if v is not None and _ENTITY_CELL_RE.match(str(v).strip()):
            last = c
    return (last + 1) if last else scan_cols + 1


def _extend_not_equal_cf(write_ws, entity_col: int, first_row: int, last_row: int) -> int:
    """Give a newly-added entity column the same red 'NOT EQUAL' highlight as the others.

    Reuses the sheet's existing containsText 'NOT EQUAL' rule fill (dxf) and adds a rule scoped to
    the new column's data range. Returns the number of rules added.
    """
    from openpyxl.formatting.rule import Rule
    from openpyxl.utils import get_column_letter

    col = get_column_letter(entity_col)
    rng = f"{col}{first_row}:{col}{last_row}"
    anchor = f"{col}{first_row}"

    seen_text = set()
    added = 0
    for cf in list(write_ws.conditional_formatting):
        for rule in cf.rules:
            if (rule.type == "containsText" and rule.dxf is not None
                    and rule.text and "NOT EQUAL" in rule.text and rule.text not in seen_text):
                seen_text.add(rule.text)
                new_rule = Rule(
                    type="containsText", operator="containsText", text=rule.text, dxf=rule.dxf,
                    formula=[f'NOT(ISERROR(SEARCH("{rule.text}",{anchor})))'],
                )
                write_ws.conditional_formatting.add(rng, new_rule)
                added += 1
    return added


def _rebuild_cf_shifted(write_ws, insert_idx: int) -> None:
    """Rebuild the sheet's conditional formatting after insert_cols(), shifting ranges to match.

    openpyxl.insert_cols() moves cell values + styles but leaves conditional-formatting ranges
    untouched, which would misalign the red 'NOT EQUAL' highlights. Columns at/after insert_idx
    shift right by one; ranges spanning the insertion point are extended by one.
    """
    from openpyxl.formatting.formatting import ConditionalFormattingList
    from openpyxl.worksheet.cell_range import CellRange

    blocks = [(cf.sqref, list(cf.rules)) for cf in write_ws.conditional_formatting]
    new = ConditionalFormattingList()
    for sqref, rules in blocks:
        parts = []
        for cr in sqref.ranges:
            mn = cr.min_col + 1 if cr.min_col >= insert_idx else cr.min_col
            mx = cr.max_col + 1 if cr.max_col >= insert_idx else cr.max_col
            parts.append(CellRange(min_col=mn, min_row=cr.min_row,
                                   max_col=mx, max_row=cr.max_row).coord)
        joined = " ".join(parts)
        for rule in rules:
            new.add(joined, rule)
    write_ws.conditional_formatting = new


def _copy_column_style(write_ws, src_col: int, dst_col: int) -> None:
    """Give a freshly-inserted (blank) column the same look as its left neighbour (styles + width)."""
    from copy import copy
    from openpyxl.utils import get_column_letter

    for r in range(1, (write_ws.max_row or 1) + 1):
        s = write_ws.cell(row=r, column=src_col)
        if not s.has_style:
            continue
        d = write_ws.cell(row=r, column=dst_col)
        d.font = copy(s.font)
        d.border = copy(s.border)
        d.fill = copy(s.fill)
        d.number_format = s.number_format
        d.protection = copy(s.protection)
        d.alignment = copy(s.alignment)
    src_letter, dst_letter = get_column_letter(src_col), get_column_letter(dst_col)
    if src_letter in write_ws.column_dimensions and write_ws.column_dimensions[src_letter].width:
        write_ws.column_dimensions[dst_letter].width = write_ws.column_dimensions[src_letter].width


def _insert_entity_in_region(write_ws, wget, entity_row: int, scan_cols: int, region: str):
    """Insert a new (blank) entity column at the END of `region`'s block; return its index, or None.

    The region header row sits directly above the entity-number row (e.g. 'Holdings' / 'AMS' /
    'EMEA' / 'APAC'). Entities are contiguous under each header, so we insert just before the next
    region's first column (or after the last entity, for the rightmost region), shifting the rest
    right and re-aligning conditional formatting + styling.
    """
    from .naming import canon_region

    region_row = entity_row - 1
    if region_row < 1:
        return None

    first_ent_col = None
    last_ent_col = 0
    for c in range(1, scan_cols + 1):
        v = wget(entity_row, c)
        if v is not None and _ENTITY_CELL_RE.match(str(v).strip()):
            if first_ent_col is None:
                first_ent_col = c
            last_ent_col = c
    if first_ent_col is None:
        return None

    headers = []                                      # (col, canonical name) for each region block
    for c in range(first_ent_col, last_ent_col + 1):
        v = wget(region_row, c)
        if v is not None and str(v).strip():
            headers.append((c, canon_region(str(v))))
    if not headers:
        return None

    target = canon_region(region)
    match_i = next((i for i, (_c, name) in enumerate(headers) if name == target), None)
    if match_i is None:
        return None

    if match_i + 1 < len(headers):
        insert_idx = headers[match_i + 1][0]          # start of the next region's block
    else:
        insert_idx = last_ent_col + 1                 # rightmost region -> after its last entity

    write_ws.insert_cols(insert_idx, 1)
    _rebuild_cf_shifted(write_ws, insert_idx)
    _copy_column_style(write_ws, insert_idx - 1, insert_idx)
    return insert_idx


# Highlight for the heading of an entity column we added (so new entities are easy to spot).
_NEW_ENTITY_FILL = PatternFill(fill_type="solid", fgColor="FFFF00")   # yellow
# Light red fill for a written cell whose source |difference| exceeds the configured threshold.
_DIFF_OVER_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")    # light red


def _to_float(v):
    """Best-effort numeric coercion of a source cell (handles 1.234,56 and '1,234.56'); None if N/A."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if not s:
        return None
    try:
        return float(s.replace(",", ""))     # drop thousands separators
    except ValueError:
        return None


def _link_entity_cell(cell, url: str) -> None:
    """Turn an entity-number cell into a hyperlink to its source file (kept blue + underlined)."""
    from copy import copy
    from openpyxl.styles import Font
    cell.hyperlink = url
    f = cell.font or Font()
    cell.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic,
                     color="0563C1", underline="single")


def _place_new_entity_column(write_ws, wget, entity_row: int, scan_cols: int,
                             entity: str, region: str | None) -> int:
    """Add a column for a new entity: under its region block when known, else at the far right.

    The entity heading cell is filled yellow so newly-added entities stand out in the workbook.
    """
    col = None
    if region:
        col = _insert_entity_in_region(write_ws, wget, entity_row, scan_cols, region)
    if col is None:
        col = _next_entity_column(wget, entity_row, scan_cols)
    cell = write_ws.cell(row=entity_row, column=col)
    cell.value = entity
    cell.fill = _NEW_ENTITY_FILL
    return col


def period_sheet_name(period: int, currency: str) -> str:
    """USD lands on 'P{period}'; other currencies on 'P{period} {CUR}' (created on demand)."""
    base = f"P{period}"
    return base if (currency or "USD").upper() == "USD" else f"{base} {currency.upper()}"


def find_template_period_sheet(wb) -> str | None:
    """Return an existing 'P<n>' sheet name to use as a structural template (or None)."""
    for name in wb.sheetnames:
        if _PSHEET_RE.match(name):
            return name
    return None


def validate_master(values_wb, year: int, period: int, master_name: str) -> list[str]:
    """Return human-readable problems (empty = OK).

    The master is a reusable template and is NOT tied to a year — the selected year only filters the
    source files. A missing P{period} sheet is fine too; it is cloned from a template at transfer
    time. So we only fail if the master has no period (P#) sheet at all to template from.
    """
    problems = []
    if find_template_period_sheet(values_wb) is None:
        problems.append("master has no period (P#) sheet to use as a template")
    return problems


def _clone_sheet(write_wb, src_name: str, target_name: str) -> None:
    """Duplicate a sheet (period or currency), best-effort copying conditional formats."""
    if target_name in write_wb.sheetnames:
        return
    src_ws = write_wb[src_name]
    new_ws = write_wb.copy_worksheet(src_ws)      # copies values, styles, dimensions, merges
    new_ws.title = target_name
    try:  # copy_worksheet does NOT carry conditional formatting; re-add the rules
        for cf in src_ws.conditional_formatting:
            for rule in cf.rules:
                new_ws.conditional_formatting.add(str(cf.sqref), rule)
    except Exception:  # best-effort; a cloned sheet without CF still holds correct values
        pass


def load_master_pair(data: bytes):
    """Load the master TWICE from the same bytes, both with data_only=True.

    Returns (values_wb, write_wb) — both hold the CACHED VALUES (not formulas):
      * values_wb -> read FORM/LINE labels and the entity row for detection.
      * write_wb  -> the copy we write into and save.

    Why data_only for the write copy? This master references an EXTERNAL workbook (30 external
    links to a SharePoint file) in its label formulas. When openpyxl re-saves those external links
    Excel flags the result as "damaged / needs repair". Loading data_only materialises every formula
    to its last-cached value, so we can drop the external links entirely and emit a self-contained
    workbook — no external references, nothing for Excel to repair. Styles and conditional
    formatting (the red NOT-EQUAL highlight) are preserved; only live formulas become static values,
    which is what a point-in-time verification copy should be anyway.
    """
    values_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    write_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    write_wb._external_links = []      # drop external-workbook links -> self-contained, no repair
    _drop_external_defined_names(write_wb)   # and the named ranges that pointed at them
    return values_wb, write_wb


_EXT_REF_RE = re.compile(r"\[\d+\]")   # external-workbook reference marker, e.g. [1]Sheet!$A$1


def _drop_external_defined_names(wb) -> None:
    """Remove named ranges that reference an external workbook (broken once links are dropped).

    Excel would otherwise report 'Removed Records: Named range' and show a repair prompt.
    Only external references (`[n]…`) are removed; internal names and table refs are left alone.
    """
    for name in list(wb.defined_names.keys()):
        if _EXT_REF_RE.search(wb.defined_names[name].value or ""):
            del wb.defined_names[name]
    for ws in wb.worksheets:               # also any worksheet-scoped names
        try:
            for name in list(ws.defined_names.keys()):
                if _EXT_REF_RE.search(ws.defined_names[name].value or ""):
                    del ws.defined_names[name]
        except Exception:
            pass


def save_workbook_to_bytes(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Shared row matching (used by BOTH transfer and verify so they always agree)
# ---------------------------------------------------------------------------

def match_source_rows(src_rows, t_keys: dict):
    """Map each non-blank source row to a target row, order-preserving per (FORM, LINE) key.

    The k-th source occurrence of a key maps to the k-th target row with that key (top-to-bottom).
    Returns a list of (form_key, line_key, check_value, abs_diff, target_row | None). Because the
    writer and verifier both call this, duplicate (FORM, LINE) lines can never be placed
    inconsistently. Rows may be 3- or 4-tuples; a missing 4th element yields abs_diff=None.
    """
    queues = {k: sorted(rows) for k, rows in t_keys.items()}
    ptr = {k: 0 for k in queues}
    out = []
    for row in src_rows:
        form, line, chk = row[0], row[1], row[2]
        diff = row[3] if len(row) > 3 else None
        if form is None and line is None:
            continue      # separator
        key = (norm_key(form), norm_key(line))
        target = None
        q = queues.get(key)
        if q is not None and ptr[key] < len(q):
            target = q[ptr[key]]
            ptr[key] += 1
        out.append((key[0], key[1], chk, diff, target))
    return out


# ---------------------------------------------------------------------------
# Transfer
# ---------------------------------------------------------------------------

def transfer_into_master(values_wb, write_wb, src: SourceData, cfg: DetectionConfig,
                         region: str | None = None) -> FileResult:
    """Write one source's check column into the matching master sheet + entity column (in memory).

    Detection/label reads come from `values_wb` (cached values); writes go to `write_wb` at the
    same coordinates so styling / conditional formatting is preserved. `region` (from the source's
    folder path) places a newly-added entity column under the matching region header block.
    """
    result = FileResult(filename=src.filename, entity=src.entity, currency=src.currency,
                        period=f"P{src.month} ({src.month}/{src.year})")
    base_name = f"P{src.month}"

    # Labels/detection come from an existing period sheet (data_only). If the requested period
    # sheet is missing, fall back to any P# sheet as the structural template and create the
    # period sheet on the write side (a clone is structurally identical, so rows/cols align).
    label_base = base_name if base_name in values_wb.sheetnames \
        else find_template_period_sheet(values_wb)
    if label_base is None:
        result.messages.append("master has no period sheet to build from -> skipped")
        return result
    if base_name not in write_wb.sheetnames:
        _clone_sheet(write_wb, label_base, base_name)

    # Currency: USD -> base sheet; otherwise 'P{period} {CUR}', cloned from the base sheet.
    target_name = period_sheet_name(src.month, src.currency)
    if target_name not in write_wb.sheetnames:
        _clone_sheet(write_wb, base_name, target_name)
    result.sheet = target_name

    ws = values_wb[label_base]
    write_ws = write_wb[target_name]
    get = _ws_getter(ws)
    max_col = min(ws.max_column or 1, 200)

    # The entity-number row is stable (insertions are column-wise). Detect it on the label sheet.
    entity_row = find_entity_header_row(get, cfg.entity_row_scan, max_col)
    if entity_row is None:
        result.messages.append(
            f"entity '{src.entity}' not found and no entity row detected in '{base_name}' -> skipped")
        return result

    # Resolve the entity's column ON THE WRITE SHEET, so it stays correct even when an earlier file
    # in this run inserted a new column (which shifts everything to its right on the write side only).
    wget = _ws_getter(write_ws)
    w_cols = min(write_ws.max_column or max_col, 400)
    found = find_entity_row_and_col(wget, cfg.entity_row_scan, w_cols, src.entity)
    if found is not None:
        _entity_row, entity_col = found
    else:
        # Entity has no column yet -> add one (under its region block when known) and note it.
        entity_col = _place_new_entity_column(write_ws, wget, entity_row, w_cols, src.entity, region)
        result.entity_added = True
        # Company name goes directly below the entity number (row 8 -> row 9 in the master).
        if src.company:
            write_ws.cell(row=entity_row + 1, column=entity_col).value = src.company
        msg = f"entity '{src.entity}' was not in '{result.sheet}' — added as new column {entity_col}"
        if region:
            msg += f" under region '{region}'"
        result.messages.append(msg)

    th = find_target_header(get, cfg.header_scan_rows, max_col, cfg)
    if th is None:
        result.messages.append(f"could not locate FORM/LINE header in '{base_name}' -> skipped")
        return result
    t_header_row, t_form_col, t_line_col = th

    # Index target rows by (FORM, LINE), ascending.
    t_first = t_header_row + 1
    t_keys: dict = {}
    last_row = ws.max_row or t_first
    for r in range(t_first, last_row + 1):
        f = norm_key(get(r, t_form_col))
        l = norm_key(get(r, t_line_col))
        if f is None and l is None:
            continue
        t_keys.setdefault((f, l), []).append(r)

    # If we added a new entity column, extend the red 'NOT EQUAL' highlight to it.
    if result.entity_added and t_keys:
        last_data_row = max(r for rows in t_keys.values() for r in rows)
        _extend_not_equal_cf(write_ws, entity_col, t_first, last_data_row)

    writes = []           # (target_row, value, over_threshold)
    for fk, lk, chk, diff, target in match_source_rows(src.rows, t_keys):
        if target is None:
            result.messages.append(
                f"line {lk!r} (form {fk!r}) not found in target -> cell skipped")
            result.skipped_lines += 1
            continue
        d = _to_float(diff)
        over = d is not None and abs(d) > cfg.diff_threshold
        writes.append((target, chk, over))

    if not writes:
        result.messages.append("no lines matched -> nothing written")
        return result

    for row, value, over in writes:
        cell = write_ws.cell(row=row, column=entity_col)
        cell.value = value
        if over:                                  # |difference| exceeds the configured threshold
            cell.fill = _DIFF_OVER_FILL
            result.highlighted += 1

    # Make the entity number (row 8) a hyperlink to its source file on SharePoint.
    if src.source_url:
        _link_entity_cell(write_ws.cell(row=entity_row, column=entity_col), src.source_url)

    result.written = len(writes)
    result.status = "done"
    return result


# ---------------------------------------------------------------------------
# Independent verification (re-reads the saved output, compares to source values)
# ---------------------------------------------------------------------------

def verify_output(output_bytes: bytes, master_bytes: bytes, sources: list[SourceData],
                  cfg: DetectionConfig, sample_limit: int = 5) -> VerifyResult:
    """Independently confirm each source line's value landed in the entity column of the output.

    Row location uses the ORIGINAL master's cached labels (labels can be formula-driven and lose
    their cache on save); the value read back comes from the SAVED output file. So the check still
    re-reads the produced artifact for the actual written values.
    """
    result = VerifyResult()
    out_wb = openpyxl.load_workbook(io.BytesIO(output_bytes), data_only=True)
    lbl_wb = openpyxl.load_workbook(io.BytesIO(master_bytes), data_only=True)
    try:
        for src in sources:
            base_name = f"P{src.month}"
            label_base = base_name if base_name in lbl_wb.sheetnames \
                else find_template_period_sheet(lbl_wb)
            target_name = period_sheet_name(src.month, src.currency)
            if label_base is None or target_name not in out_wb.sheetnames:
                continue
            out_ws = out_wb[target_name]        # values were written here (period/currency sheet)
            lbl_ws = lbl_wb[label_base]         # labels read from an existing template period sheet
            lget = _ws_getter(lbl_ws)
            oget = _ws_getter(out_ws)
            max_col = min(lbl_ws.max_column or 1, 200)
            out_cols = min(out_ws.max_column or max_col, 400)

            # Entity column comes from the OUTPUT (it includes any column we had to add).
            ent = find_entity_row_and_col(oget, cfg.entity_row_scan, out_cols, src.entity)
            th = find_target_header(lget, cfg.header_scan_rows, max_col, cfg)
            if ent is None or th is None:
                continue
            _erow, ecol = ent
            t_header_row, t_form_col, t_line_col = th

            t_keys: dict = {}
            last_row = lbl_ws.max_row or t_header_row
            for r in range(t_header_row + 1, last_row + 1):
                f = norm_key(lget(r, t_form_col))
                l = norm_key(lget(r, t_line_col))
                if f is None and l is None:
                    continue
                t_keys.setdefault((f, l), []).append(r)

            for fk, lk, chk, diff, target in match_source_rows(src.rows, t_keys):
                if target is None:
                    continue      # unmatched line (writer also skips these)
                got = out_ws.cell(row=target, column=ecol).value
                result.checked += 1
                if norm(got) != norm(chk):
                    result.mismatches += 1
                    if len(result.samples) < sample_limit:
                        cell = out_ws.cell(row=target, column=ecol).coordinate
                        result.samples.append(
                            f"{target_name}!{cell} line {lk!r}: expected {chk!r} got {got!r}")
        return result
    finally:
        out_wb.close()
        lbl_wb.close()
