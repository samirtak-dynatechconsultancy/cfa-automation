# CFA Automation — FastAPI Application Build Plan

**Status:** Plan / design. No application code written yet.
**Supersedes** the trigger/UI open questions in `INFRA_REQUIREMENTS.md`.

---

## 1. Decisions locked (from Q&A)

| Area | Decision |
|---|---|
| Framework | **FastAPI**, server-rendered UI (**Jinja2 templates + light vanilla JS**). One deployable. |
| Excel engine | **Pure-Python (openpyxl)** — no desktop Excel / xlwings. Runs anywhere Python runs. |
| Run target | **Local now** (each user runs it, browses to `localhost`), **architected to deploy to Azure later** with minimal change. |
| SharePoint access | **Microsoft Graph, app-only** (client credentials). **Tenant-wide browse** of all sites. |
| Graph permission | `Sites.ReadWrite.All` (application) — covers browsing all sites (read) **and** writing output. Admin consent required. |
| Credential storage | **`.env` file** (client id / tenant id / client secret). No Key Vault for now. |
| Per-run inputs | **User picks everything each run** — source folder, master workbook, output folder — via a SharePoint browser. Defaults pre-fill from settings but are overridable. |
| Output | **Timestamped copy** (`… _autofilled_<ts>.xlsx`) into a **configurable output folder**. Master never overwritten. |
| Settings/defaults | Stored in a **`settings.xlsx` workbook kept in SharePoint**, read/written via Graph. No database. |
| UI auth | **None** for now (local tool). Flagged as a must-revisit before any central hosting. |

---

## 2. What we reuse from the existing code

The existing script's logic is **library-agnostic in the right places** — the detection/parsing helpers take a `get_cell(r, c)` callable, so they work unchanged whether the values come from xlwings or openpyxl:

- **Reuse as-is:** `norm`, `norm_key`, `parse_entity`, `parse_period`, `find_header_row`,
  `find_target_header`, `find_entity_row_and_col`. (These are already used by *both* `cfa_transfer.py`
  and `verify_transfer.py`.)
- **Port from xlwings → openpyxl:** `read_source` (block reads) and `transfer_to_target` (the
  position-transfer + label-verification + graceful-skip + contiguous-write logic). The **algorithm is
  kept identical**; only the cell-access layer changes. `verify_transfer.py` is **already openpyxl**, so
  it becomes the in-process verification step almost unchanged.

### ⚠️ The one real engineering risk — formatting fidelity
xlwings was chosen originally to preserve the **red NOT-EQUAL conditional-formatting highlights**.
Moving to openpyxl:
- **Conditional-formatting *rules* live in the master template**, and openpyxl **preserves existing CF
  rules** across load→save. Because the highlight is driven by a rule (`= "NOT EQUAL"`), writing the
  value re-fires the rule when a human opens the file — so this **should** carry over.
- **BUT** openpyxl can drop other workbook features on round-trip (charts, some styles, VBA, form
  controls, pivot caches). **Mandatory early spike:** load the real master with openpyxl, write a
  column, save, and **diff against a known-good** to confirm the highlights and everything else the
  client cares about survive. This spike gates the whole approach — do it in Phase 0.

---

## 3. Architecture

```
Browser (Jinja2 pages + fetch/JS)
   │  HTTP
   ▼
FastAPI app  ──────────────────────────────────────────────
   ├── UI routes (Jinja2): / (run page), /admin
   ├── API routes (JSON):
   │      /api/sites          browse: search/list sites
   │      /api/drives         list document libraries for a site
   │      /api/items          list folders/files under a drive/folder
   │      /api/runs (POST)    start a run (async) → run_id
   │      /api/runs/{id}      poll status/result
   │      /api/settings       GET/PUT admin defaults
   ├── Graph client (httpx + MSAL app-only token, cached)
   ├── Processing core (openpyxl port of transfer + verify)
   └── Settings service (reads/writes settings.xlsx via Graph)
        │
        ▼
   Microsoft Graph  →  SharePoint (sources, master, output, settings.xlsx)
```

**Request flow for a run (async):**
1. User browses SharePoint in the UI, selects **source folder**, **master workbook**, **output folder**.
2. `POST /api/runs` validates input, creates a `run_id`, kicks off a **background task**, returns immediately.
3. Background task: download master + source files from Graph → temp dir → run the (openpyxl) transfer →
   run verify → upload timestamped output to the output folder via Graph → record status/summary.
4. UI **polls `/api/runs/{id}`** and shows live status → final PASS/FAIL summary + link to the output file.

---

## 4. SharePoint / Graph integration

- **Auth:** app-only **client credentials** (MSAL `ConfidentialClientApplication` or `azure-identity`
  `ClientSecretCredential`), token scope `https://graph.microsoft.com/.default`, token cached and
  refreshed in-process.
- **Browse all sites:** site search endpoint (search box in UI) + list document libraries (drives) +
  navigate folders via driveItem children. (Listing *every* site blindly is not ideal UX; a **search-as-
  you-type site picker** is the plan.)
- **Read files:** download item content to a temp working dir.
- **Write output:** upload the timestamped copy to the chosen output folder (upload session to be safe
  for larger files).
- **Settings:** `settings.xlsx` at a **fixed, configured SharePoint path** — read on startup/admin load,
  written on admin save.
- **Resilience carried over from the script:** per-file try/except (one bad file can't kill the batch),
  DONE/SKIP/ERR summary surfaced in the UI, graceful per-line skip on template drift.

---

## 5. UI pages

- **Run page (`/`):** SharePoint browser (site → library → folder), three pickers (source folder, master,
  output folder) pre-filled from defaults, a **Run** button, and a **live status/result panel**
  (DONE/SKIP per file, verify PASS/FAIL, link to output).
- **Admin page (`/admin`):** form to view/edit defaults — default site, default source folder, default
  master, default output folder, and the detection config (source tab name, FORM/LINE/EQUAL-TO header
  anchors, scan bounds). **Save** writes back to `settings.xlsx` in SharePoint.
  *(No auth gate for now — see §8 risk.)*

---

## 6. Project structure (proposed)

```
app/
  main.py              FastAPI app, route registration
  config.py            env settings (pydantic-settings)
  graph/               token provider + Graph client (sites/drives/items/content)
  core/                openpyxl port of read_source / transfer_to_target + verify
  services/            run manager (async jobs + status), settings service (settings.xlsx)
  routers/             ui.py, sites.py, runs.py, settings.py
  templates/           Jinja2: base, run, admin
  static/              light JS + CSS
tests/
.env                   (git-ignored) Graph credentials + config
requirements.txt       + fastapi, uvicorn, jinja2, python-multipart, httpx, msal (or azure-identity), pydantic-settings, python-dotenv
```
(Existing `cfa_transfer.py` / `verify_transfer.py` remain the reference; shared helpers get lifted into `app/core`.)

---

## 7. Configuration / `.env` (names only — no values)

- `GRAPH_TENANT_ID`
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`
- `SETTINGS_WORKBOOK_PATH` — SharePoint location of `settings.xlsx` (site + drive + path)
- `LOG_LEVEL`
- `APP_PORT` / host bind
- *(detection defaults — SOURCE_TAB, FORM/LINE/CHECK header anchors, scan bounds — live in `settings.xlsx`, editable via the admin page, with code fallbacks)*

---

## 8. Risks & things to revisit

1. **Formatting fidelity (Phase 0 gate)** — validate openpyxl preserves the red CF highlights on the
   real master before committing. *Highest risk.*
2. **No UI auth + `Sites.ReadWrite.All`** — a local, unauthenticated app holding a **tenant-wide
   read/write** Graph credential is powerful. Fine for a controlled local prototype; **must add sign-in
   and tighten the permission (e.g. `Sites.Selected`) before any shared/central hosting.**
3. **`.env` secret** — acceptable locally; becomes Key Vault / managed identity when hosted.
4. **In-process background jobs** — fine for local single-user; won't survive scale-out/restart when
   hosted → would move to a queue/Container Apps Job. Keep the run-manager behind an interface so this
   swaps cleanly.
5. **Concurrent settings writes** — `settings.xlsx` in SharePoint has no locking; low risk with one
   admin, but note it.

---

## 9. Build phases

- **Phase 0 — Formatting spike (gate):** openpyxl load→write→save on the real master; confirm highlights
  and workbook integrity. Decide go/no-go on pure-Python.
- **Phase 1 — Core port:** move shared helpers to `app/core`; reimplement `read_source` /
  `transfer_to_target` on openpyxl; wire in verify. Prove parity with the current script on the sample files.
- **Phase 2 — Graph layer:** app-only token + client; sites/drives/items browse; download/upload.
- **Phase 3 — API + run manager:** `/api/runs` async start + status polling; per-file summary.
- **Phase 4 — UI:** run page with SharePoint browser + result panel.
- **Phase 5 — Admin + settings.xlsx** in SharePoint (load defaults, save defaults).
- **Phase 6 — Hardening:** logging, error surfacing, docs; note the Azure-later deltas (auth, Key Vault,
  job queue).

---

## 10. Remaining questions (non-blocking — sensible defaults assumed)

1. **App details** — please provide `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`, and
   confirm **admin consent** for `Sites.ReadWrite.All` is granted.
2. **`settings.xlsx` location** — which SharePoint site/library/path should hold it? (I'll create the
   workbook structure.)
3. **Verify step** — assumed **always run** after transfer and shown as PASS/FAIL in the UI. OK?
4. **Output delivery** — assumed UI shows a **link to the output in SharePoint** (plus optional download).
   Do you also want it emailed / posted to Teams? (assumed **no** for now.)
5. **Multiple source files per run** — assumed the run processes **all entity files in the chosen source
   folder** (as the script does today), master auto-routing by period/entity from file contents. OK?
6. **Concurrency** — assumed **one run at a time per user**, single-user local. OK?
7. **Python version** — standardize on **3.12** (repo venv is 3.13; both fine). Any constraint?
