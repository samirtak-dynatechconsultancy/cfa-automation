# CFA Consistency-Check Transfer

Automates the manual copy job: for every entity source workbook, the consistency-check results
(`OK` / `NOT EQUAL` / `-----`) from the consistency-check tab (**G**, column **D**, rows ~10–210)
are written into the matching entity **column** of the shared `CFA verification_Pn YYYY.xlsx`
master, on the correct period sheet (`P5`, `P4`, …).

It uses **xlwings**, which drives the installed desktop Excel, so all formatting — including the red
conditional-formatting highlights on `NOT EQUAL` — is preserved. (Pure `openpyxl` would strip them.)

## What it does each run

1. Looks in one **folder** that contains *both* the source files and the target master.
2. Identifies the **target** = the `.xlsx` whose name contains `verification` (errors out if it can't
   pick exactly one), and excludes it + Excel lock files (`~$…`) + previous `_autofilled_` copies
   from the source set.
3. For each remaining `.xlsx`, reads the **entity number** and **period** from *inside* the file
   (`A2` = `NUMBER: 055`, `A3` = `PERIOD: 5 / 2026`) — never from the filename.
4. Routes each source: **period → `P{month}` sheet**, **entity number → column** (matched against
   row 8, trailing spaces stripped).
5. Transfers the check column line-by-line by **row position with per-row label verification**
   (correct even for the few duplicate `(FORM, LINE)` lines; logs and skips a line instead of
   mis-pasting if a template line was added/removed/renamed).
6. Saves a **timestamped copy** next to the master (`… _autofilled_YYYYMMDD_HHMMSS.xlsx`) — the
   original master is never modified.
7. Prints a **summary**: files done / skipped (with reasons) and any per-line warnings.

Missing period sheet or entity column → that file is **skipped and logged**, the run continues.

## Requirements

- Windows with **Microsoft Excel** installed (xlwings drives it).
- Python 3.10+ and the packages in `requirements.txt`.

```powershell
py -3.12 -m pip install -r requirements.txt
```

> SharePoint note: point the script at the **locally synced** folder (OneDrive/SharePoint sync),
> not a web URL. No Excel-online dependency.

## Usage

Run against the folder the script lives in (default):

```powershell
py -3.12 cfa_transfer.py
```

Or point it at another folder / show Excel while it runs:

```powershell
py -3.12 cfa_transfer.py --folder "C:\Users\me\OneDrive\... \CFA P5" --visible
```

### Verify the result (independent check)

`verify_transfer.py` re-reads the saved copy with **openpyxl** (a different library than the writer)
and asserts every written cell equals the source value for that `(FORM, LINE)` line.

```powershell
py -3.12 verify_transfer.py          # auto-picks the newest *_autofilled_*.xlsx in the folder
py -3.12 verify_transfer.py "C:\...\CFA verification_P5 2026 _autofilled_20260623_101500.xlsx"
```

Exit code `0` = all pass, `1` = mismatches found.

## Configuration

Edit the `CONFIG` block at the top of `cfa_transfer.py`:

| Key | Meaning | Default |
|-----|---------|---------|
| `FOLDER` | Folder with sources + target | script's own folder |
| `TARGET_NAME_TOKEN` | Substring identifying the master | `"verification"` |
| `SOURCE_TAB` | Preferred consistency-check tab (auto-detected if renamed) | `"G"` |
| `FORM_HEADER` / `LINE_HEADER` / `CHECK_HEADER` | Header anchor strings used to locate the table | `FORM` / `LINE` / `EQUAL TO` |
| `EXCEL_VISIBLE` | Show Excel while running | `False` |

Because sheet, header row, and columns are located by **header anchors** (not fixed cell
references), a renamed tab, an inserted line, or a shifted column does not break the run.

## Files

- `cfa_transfer.py` — main tool.
- `verify_transfer.py` — independent post-run verification.
- `requirements.txt` — dependencies.
- `README.md` — this file.
