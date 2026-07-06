"""Reorganize the messy CFA folder into the clean convention.

Reads a source root (default '01 - CFA - Hardcoded') and produces a clean tree:

    <OUT>/<YYYY>/<REGION>/<ENTITY [- Company]>/P<PP>/<original filename>

Rules
- A file is a candidate CFA source if its name starts with 'CFA <YYPP> ...' (4 digits, no dash),
  i.e. the reported/export workbooks. Input/View macro files ('CFA - 2025_P..'), 'CFA light',
  'TFA', Flash/Dashboard/Financials/Notes, PDFs and DOCs are ignored.
- Year + period come from the FILENAME token (CFA 2605 -> 2026 / P05) — reliable regardless of the
  folder name variants.
- Region is the folder directly under the year (AMS/APAC/EMEA/HOLDINGS); files not under a known
  region go to '_UNSORTED' (nothing is lost).
- The entity folder uses one canonical '<code> - <Company>' name per entity (so 265 and
  '265 - TenCate Canada' don't split); files sit in a 'P<PP>' subfolder under it.
- Folders named OLD/old/archive/_archive are skipped.
- MULTIPLES ARE KEPT: every source file for a period is copied (original name preserved), so
  duplicates/versions (V2, (2), dates, Final/POST AUDIT) all survive. The app picks the latest at
  run time; the folder retains everything.
- Currency comes from the name (USD/EUR/EGP/MYR/...); when absent it defaults to USD.

Usage:  py reorganize_cfa.py                 # dry run (prints the plan + stats)
        py reorganize_cfa.py --apply         # actually copy into the output folder
        py reorganize_cfa.py --src "..." --out "..."
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
from collections import defaultdict

KNOWN_REGIONS = {"AMS", "APAC", "EMEA", "HOLDING", "HOLDINGS"}
REGION_CANON = {"HOLDING": "HOLDINGS"}
IGNORE_DIRS = {"old", "archive", "_archive", "olds", "bak"}
CURRENCIES = {"USD", "EUR", "EGP", "MYR", "GBP", "CAD", "CHF", "JPY", "CNY", "INR",
              "AUD", "SGD", "HKD", "MXN", "BRL", "ZAR", "SEK", "NOK", "DKK", "PLN", "TRY", "AED"}

# CFA followed by a 4-digit YYPP token (no dash after CFA => excludes 'CFA - 2025_P..').
CANDIDATE_RE = re.compile(r"^CFA\s+(\d{2})(\d{2})\b", re.IGNORECASE)
ENTITY_RE = re.compile(r"\b(\d{3}[A-Za-z]?)\b")


def parse_name(fname: str):
    """Return (year, period, entity, currency) or None if not a CFA source file."""
    stem = os.path.splitext(fname)[0]
    if "light" in stem.lower():
        return None
    m = CANDIDATE_RE.match(stem)
    if not m:
        return None
    yy, pp = m.group(1), m.group(2)
    year = 2000 + int(yy)
    period = int(pp)
    rest = stem[m.end():]                       # after 'CFA 2605'
    ent = ENTITY_RE.search(rest)                # first 3-digit(+letter) token = entity
    if not ent:
        return None
    entity = ent.group(1).upper()
    cur = None
    for tok in re.findall(r"\b([A-Za-z]{3})\b", rest):
        if tok.upper() in CURRENCIES:
            cur = tok.upper()
            break
    return year, period, entity, cur or "USD"


def region_of(rel_parts: list[str]) -> str:
    """rel_parts is the path under the source root: [YEAR, REGION?, ...]. Return canonical region."""
    if len(rel_parts) >= 2:
        cand = rel_parts[1].strip().upper()
        if cand in KNOWN_REGIONS:
            return REGION_CANON.get(cand, cand)
    return "_UNSORTED"


def entity_folder(entity: str, rel_parts: list[str]) -> str:
    """Prefer an '<entity> - <Company>' folder name found in the source path; else the entity code."""
    for part in rel_parts:
        p = part.strip()
        if re.match(rf"^{re.escape(entity)}\b", p, re.IGNORECASE) and len(p) > len(entity):
            return re.sub(r"\s*[-–]\s*", " - ", p)      # normalise the dash spacing
    return entity


def collect(src_root: str):
    items: list[dict] = []          # every source-candidate file (no dedup)
    # One canonical folder name per (year, region, entity): if ANY file for that entity sits in a
    # '<entity> - <Company>' folder, use that name for ALL of the entity's files (so 265 and
    # '265 - TenCate Canada' don't become two folders).
    entity_names: dict[tuple, str] = {}
    stats = defaultdict(int)
    for dirpath, dirnames, filenames in os.walk(src_root):
        dirnames[:] = [d for d in dirnames if d.strip().lower() not in IGNORE_DIRS]
        rel = os.path.relpath(dirpath, src_root)
        rel_parts = [] if rel == "." else rel.split(os.sep)
        for fn in filenames:
            if not fn.lower().endswith((".xlsx", ".xlsm")):
                continue
            parsed = parse_name(fn)
            if parsed is None:
                stats["ignored_nonsource"] += 1
                continue
            year, period, entity, cur = parsed
            region = region_of(rel_parts)
            items.append({"path": os.path.join(dirpath, fn), "name": fn, "region": region,
                          "year": year, "period": period, "entity": entity, "cur": cur})
            stats["candidates"] += 1
            nm = entity_folder(entity, rel_parts)
            if nm != entity:
                nk = (year, region, entity)
                if nk not in entity_names or len(nm) > len(entity_names[nk]):
                    entity_names[nk] = nm
    return items, entity_names, stats


def plan(items: list, entity_names: dict):
    rows = []
    seen: set = set()
    for it in sorted(items, key=lambda x: (x["year"], x["region"], x["entity"],
                                           x["period"], x["name"])):
        ent_folder = entity_names.get((it["year"], it["region"], it["entity"]), it["entity"])
        dest_dir = os.path.join(str(it["year"]), it["region"], ent_folder, f"P{it['period']:02d}")
        dest = os.path.join(dest_dir, it["name"])
        base, ext = os.path.splitext(it["name"])
        n = 1
        while dest.lower() in seen:            # avoid clobbering a same-named file from elsewhere
            n += 1
            dest = os.path.join(dest_dir, f"{base} ({n}){ext}")
        seen.add(dest.lower())
        rows.append((dest, it["path"]))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default="01 - CFA - Hardcoded")
    ap.add_argument("--out", default="02 - CFA - Reorganized")
    ap.add_argument("--apply", action="store_true", help="actually copy (default: dry run)")
    args = ap.parse_args()

    items, entity_names, stats = collect(args.src)
    rows = plan(items, entity_names)

    print(f"SOURCE: {args.src}")
    print(f"OUTPUT: {args.out}   ({'APPLY' if args.apply else 'DRY RUN'})\n")
    regions = defaultdict(int)
    years = defaultdict(int)
    groups = defaultdict(int)      # (year,region,entity,period,cur) -> count, to spot multiples
    for it in items:
        groups[(it["year"], it["region"], it["entity"], it["period"], it["cur"])] += 1
    for out_rel, _ in rows:
        parts = out_rel.split(os.sep)
        years[parts[0]] += 1
        regions[parts[1]] += 1
    multiples = {k: n for k, n in groups.items() if n > 1}
    print(f"Candidates seen : {stats['candidates']}")
    print(f"Non-source files ignored: {stats['ignored_nonsource']}")
    print(f"Files kept (multiples preserved): {len(rows)}")
    print(f"Period slots with multiple files: {len(multiples)}")
    print(f"By year   : {dict(sorted(years.items()))}")
    print(f"By region : {dict(sorted(regions.items()))}\n")
    if multiples:
        print("Examples of kept multiples (entity/period -> count):")
        for (y, rg, ent, per, cur), n in list(sorted(multiples.items()))[:6]:
            print(f"  {y} {rg} {ent} P{per:02d} {cur}: {n} files")
        print()

    print("Sample of planned files (first 25):")
    for out_rel, srcp in rows[:25]:
        print(f"  {out_rel}")

    if args.apply:
        if os.path.isdir(args.out):
            shutil.rmtree(args.out)          # rebuild cleanly (drops any prior layout)
        copied = 0
        for out_rel, srcp in rows:
            dest = os.path.join(args.out, out_rel)
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(srcp, dest)
            copied += 1
        print(f"\nCopied {copied} file(s) into '{args.out}'.")
    else:
        print(f"\n(dry run — re-run with --apply to copy {len(rows)} files)")


if __name__ == "__main__":
    main()
