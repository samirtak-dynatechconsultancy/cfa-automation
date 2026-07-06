# CFA Consistency-Check Transfer — Web App

FastAPI app that transfers entity consistency-check results into a shared verification master,
sourcing files from **SharePoint via Microsoft Graph** (app-only). Server-rendered UI: browse
SharePoint (site → library → folder), pick the source folder / master / output folder, run, and
see an independent PASS/FAIL verification. The master is never modified — a timestamped copy is
written to the output folder.

Pure-Python engine (**openpyxl**), no desktop Excel — runs anywhere. Ported from the original
`cfa_transfer.py` / `verify_transfer.py`; the detection/parsing logic is shared unchanged.

## Layout

```
app/
  main.py            FastAPI app + router registration
  config.py          env/.env settings
  deps.py            singleton wiring
  core/              detection.py, engine.py (openpyxl transfer+verify), models.py
  graph/             auth.py (MSAL app-only token), client.py (Graph REST)
  services/          runs.py (async run manager), settings_store.py (SharePoint List, auto-created)
  routers/           ui.py, sites.py, runs.py, settings.py
  templates/         base / run / admin (Jinja2)
  static/            app.js, style.css
run.py               local launcher
requirements.txt
.env.example
```

## Setup

```powershell
# from cfa_app\
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
copy .env.example .env      # then fill in GRAPH_* values
.\.venv\Scripts\python.exe run.py
# open http://127.0.0.1:8000
```

## Microsoft Entra app registration (what the client provisions)

- App registration (single tenant).
- **Microsoft Graph application permission** `Sites.ReadWrite.All` (browse all sites + write output),
  with **admin consent**. (Least-privilege alternative: `Sites.Selected` granted per-site — narrows
  browsing to registered sites.)
- **Auto-creating the settings List** additionally needs `Sites.Manage.All` (creating a list is a
  manage operation). If you'd rather not grant that, create the `CFA App Settings` list once by hand
  (columns: Title, Value) — then `Sites.ReadWrite.All` is enough to read/write its items.
- A **client secret** (put its value in `.env` as `GRAPH_CLIENT_SECRET`).

## Endpoints

- `GET /` — run page · `GET /admin` — defaults/detection
- `GET /api/sites?q=` · `GET /api/drives?site_id=` · `GET /api/items?drive_id=&item_id=`
- `POST /api/runs` → `{run_id}` · `GET /api/runs/{run_id}` — status/result
- `GET /api/settings` · `PUT /api/settings`
- `GET /health`

## Notes / current scope

- **No UI authentication** (local tool). Add Entra sign-in + tighten the Graph permission before any
  shared/central hosting.
- **Secrets in `.env`**; move to Key Vault / managed identity when hosted.
- **In-process background runs** — fine locally; swap for a queue/job when hosted (contained in
  `services/runs.py`).
- **Conditional formatting:** standard rules (the red NOT-EQUAL highlight) survive the openpyxl
  round-trip; the 2010+ CF *extension* (data bars / icon sets) is dropped by openpyxl — visually
  confirm on the master if it uses those.
```
