# Deploy to PythonAnywhere

FastAPI (ASGI) runs on PythonAnywhere's WSGI hosting via the `a2wsgi` bridge (`wsgi.py`).

## Prerequisites
- A **paid plan (Hacker $5/mo or higher)** — required so the app can make **outbound** calls to
  `graph.microsoft.com` and `login.microsoftonline.com`. The free tier blocks non‑allowlisted
  outbound traffic and Graph will fail.
- Your Microsoft Entra app details (`GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`) and
  `SETTINGS_SITE_URL`.

## 1. Get the code (Bash console on PythonAnywhere)
```bash
git clone https://github.com/samirtak-dynatechconsultancy/cfa-automation.git
cd cfa-automation/cfa_app
```

## 2. Create a virtualenv and install deps
```bash
mkvirtualenv --python=/usr/bin/python3.11 cfa
pip install -r requirements.txt
```
(`mkvirtualenv` also activates it; note its path, e.g. `/home/USERNAME/.virtualenvs/cfa`.)

## 3. Create the .env (in cfa-automation/cfa_app)
```bash
cp .env.example .env
nano .env   # fill in GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET / SETTINGS_SITE_URL
```

## 4. Create the web app
- **Web** tab → **Add a new web app** → **Manual configuration** → **Python 3.11**.
- **Virtualenv:** set it to `/home/USERNAME/.virtualenvs/cfa`.
- **Source code / working directory:** `/home/USERNAME/cfa-automation/cfa_app`.

## 5. Point the WSGI file at the app
Edit the WSGI configuration file (the link is on the Web tab, e.g.
`/var/www/USERNAME_pythonanywhere_com_wsgi.py`). Replace its contents with:

```python
import os
import sys

project = "/home/USERNAME/cfa-automation/cfa_app"
if project not in sys.path:
    sys.path.insert(0, project)
os.chdir(project)          # so .env loads and relative paths resolve

from wsgi import application   # noqa: E402  (a2wsgi-wrapped FastAPI app)
```
Replace `USERNAME` everywhere with your PythonAnywhere username.

## 6. Reload
Click the big green **Reload** button on the Web tab. Open
`https://USERNAME.pythonanywhere.com/` — the Run page should load.

## Notes / limits
- **One web worker.** Keep the default single worker so the in‑memory run registry and the
  live‑progress polling stay consistent (a run started on one worker must be polled on the same one).
- **Long runs.** A single HTTP request is capped (~5 min) but runs execute in a **background thread**
  (the POST returns immediately), so that cap doesn't apply to the run itself. Very large scopes still
  depend on the worker staying alive; prefer narrower scopes (a region + period range).
- **Updating:** `git pull` in the console, then **Reload** on the Web tab.
- **Secrets:** `.env` is git‑ignored — it lives only on the server; never commit it.
