# CFA Consistency-Check Automation — Azure Infrastructure Requirements

**Prepared for:** Client provisioning / cloud onboarding
**Prepared by:** Solutions Architecture
**Status:** Requirements & analysis only — no code, config, or IaC included.

---

## 0. Critical finding (read first)

The brief describes the application **we intend to build**: a FastAPI HTTP service that reads
documents from SharePoint via Microsoft Graph using Entra app-only auth. **The code in this
workspace is not that application yet.** Evidence:

| Brief says | Codebase actually contains | Evidence |
|---|---|---|
| FastAPI / ASGI HTTP service | A **batch command-line script** (run-to-completion), no web server | `cfa_transfer.py` `main()` uses `argparse`; no `fastapi`/`uvicorn`/`starlette` anywhere in app code |
| A **UI** for users to enter SharePoint paths and trigger runs | **No UI, no HTTP layer** — inputs come from a `CONFIG` dict + CLI flags | `CONFIG = {...}` at top of `cfa_transfer.py`; `argparse` flags `--folder/--token/--visible`; no frontend assets, templates, or web framework |
| Reads from SharePoint via Microsoft Graph | Reads from a **local/synced filesystem folder** | `discover_files()` uses `glob.glob(os.path.join(folder, "*.xlsx"))`; README: *"point the script at the locally synced folder … not a web URL"* |
| Entra app-only auth (MSAL/Graph) | **No auth code at all** | No `msal`, `azure-*`, `requests`, `httpx`, or Graph calls in `requirements.txt` or source |
| Cloud-native (Linux/serverless) | **Windows + desktop Excel automation (COM)** | `import xlwings as xw`; `xw.App(...)` drives installed Excel; `pywin32`/`win32com` present in the venv; README: *"Windows with Microsoft Excel installed"* |

**Consequence for infrastructure:** the current write engine, **xlwings, automates the desktop
Excel application through COM.** Server-side Office automation is **not supported on Azure PaaS**
(App Service sandbox, Container Apps/Functions on Linux) — it requires a Windows OS with Microsoft
Office installed. **This one dependency, left as-is, forces IaaS (a Windows VM with Office) and
rules out every serverless/PaaS option.** The hosting recommendation below therefore hinges on a
decision the client must make (see §2 and §6).

Everything below is grounded in the code as it exists, with each **inferred / target-state** item
explicitly flagged.

---

## 1. Summary — what the app does and its execution model

**Function (from code):** For every "entity" source workbook in a folder, it reads consistency-check
results (`OK` / `NOT EQUAL` / `-----`) from a check tab and writes them into the matching entity
*column* of a shared master workbook (`CFA verification_Pn YYYY.xlsx`), on the correct period sheet
(`P5`, `P4`, …). It saves a **timestamped copy** and never modifies the original master.
A second script, `verify_transfer.py`, independently re-reads the output with a different library
(openpyxl) and asserts every written value matches its source.

**Execution model (from code):**
- **Batch / run-to-completion**, invoked from the command line (`cfa_transfer.py`, `verify_transfer.py`).
- **No HTTP server, no UI, no background workers, no in-code scheduler.** The "FastAPI service,
  endpoints, background tasks" model — **and the user-facing UI for entering SharePoint paths and
  triggering runs** — in the brief is **target-state and not present**.
- Processes files **sequentially**, one open Excel instance for the whole run.

**Target-state execution model (from the brief — inferred, to be built):**
- A **persistent FastAPI/ASGI service** exposing endpoints to **start a run** and **poll its
  status/result**, plus a **web UI** where a user supplies the SharePoint site/library/folder paths
  and other run parameters and presses "run".
- Because a run is long-running (see resource profile), the trigger endpoint should **accept the
  request and process asynchronously** (background task / job), with the UI polling for completion —
  **not** a synchronous request that blocks for the whole run.
- The UI makes the trigger **interactive/manual and user-initiated**, which resolves the earlier
  "how is a run triggered" question toward a **human-driven, authenticated web front end** (a
  scheduler is now optional, not required).

**Runtime & key dependencies (from code):**
- **Python** — venv is 3.13; `README` says 3.10+. Requirements pin: `xlwings>=0.30`, `openpyxl>=3.1`.
- **xlwings** (writer) → requires **desktop Microsoft Excel** + **pywin32/COM** on **Windows**.
- **openpyxl** (independent verifier) → pure Python, cross-platform.
- No networking, database, or cloud SDK dependencies are declared.

**External services (from code):** **None over the network.** SharePoint is used only as a
**locally synced folder** today (per README). Microsoft Graph / SharePoint API integration is
**target-state, not implemented.**

**Resource profile (from code + reasoning, flagged inferred):**
- CPU/memory is dominated by the **Excel process** and workbook size, not the Python logic
  (*inferred* — no profiling in repo). The scripts cap scans (`MAX_DATA_ROWS = 400`,
  `HEADER_SCAN_ROWS = 20`) and batch writes into contiguous ranges, so the Python side is light.
- **Concurrency: 1.** One `xw.App` drives one Excel; files handled in a sequential loop. Not designed
  for parallel requests.
- **Long-running:** run time scales with the number of source files and Excel COM round-trips
  (*inferred* — seconds-to-minutes per file). Any HTTP wrapper must treat a run as a **long async job**,
  not a synchronous request.

**State it keeps (from code):**
- **No database, no cache.** Only output is a **new `.xlsx` file** written next to the master
  (`… _autofilled_YYYYMMDD_HHMMSS.xlsx`). Everything else is in-memory for the run.
- Reads/writes are **whole-workbook file operations**; scratch space is ephemeral.

**Secrets & configuration (from code):**
- **None today.** Configuration is a `CONFIG` dict at the top of `cfa_transfer.py` plus CLI flags
  (`--folder`, `--token`, `--visible`). **No environment variables and no secrets** are read.
- The target (Graph app-only auth) will introduce the first real secrets — see §4/§5.

**Networking (from code):** **None.** Purely local filesystem. Public/private/VNet needs are entirely
**target-state**, introduced only when Graph + an HTTP endpoint are added.

---

## 2. Recommended compute host

The recommendation splits on the **Excel-engine decision**, because it, not FastAPI, is the binding
constraint.

### Recommendation A (preferred, cloud-native): re-platform the write engine, then use **Azure Container Apps**

**Precondition:** move the write path off desktop Excel (xlwings) to a **pure-Python engine
(openpyxl)** so the service runs on Linux. *Trade-off to state plainly:* openpyxl does **not** carry
over the red conditional-formatting highlights automatically the way desktop Excel does — that
behaviour would have to be reproduced in code. This is the key functional trade of going cloud-native.

**Why Container Apps for the target FastAPI service:**
- Purpose-built for a **persistent HTTP/ASGI service** (FastAPI/uvicorn) — the brief's execution model.
- **Serves the UI too:** the same container can host the web UI (FastAPI serving a static/SPA front
  end or server-rendered pages) alongside the API, so no separate frontend host is strictly required.
  *(Alternative: host the UI on **Azure Static Web Apps** and call the API — only worth it if the
  frontend is a standalone SPA the client wants deployed independently; not required by anything in
  the code.)*
- **Built-in authentication** ("Easy Auth"-style, Entra ID) can gate the UI/endpoints at the platform
  level, so end users must sign in before they can trigger a run (see §3, Group 1b).
- **Scale-to-zero** and scale-out by HTTP concurrency — fits a workload that is idle between
  month-end runs but must accept an inbound, user-initiated trigger.
- First-class **managed identity** support → lets the app get Graph app-only tokens **without storing
  a secret** (see §3).
- Supports the **long-running job** shape: expose an endpoint that accepts a request and processes in
  the background (UI polls for status), or split heavy processing into a **Container Apps Job**.

**When to pick an alternative instead:**
- **Azure App Service (Linux, Web App):** choose if the client is standardized on App Service and
  doesn't need scale-to-zero. Equally valid for a single persistent FastAPI container; slightly simpler,
  no separate registry/environment concepts. **Not** viable with xlwings (sandbox blocks Office automation).
- **Azure Functions:** only if the design is genuinely event/timer-driven and short per-invocation.
  A persistent FastAPI service is a poor fit for the Functions model, and Functions still cannot run
  desktop Excel. **Not recommended** given the brief explicitly calls for a long-running ASGI service.

### Recommendation B (only if the red-highlight formatting MUST be preserved via desktop Excel)

Keep xlwings → you **must** host on a **Windows VM (or VM Scale Set) with Microsoft Office installed**
(IaaS). FastAPI runs there under a process manager. This forfeits scale-to-zero, PaaS patching, and
serverless economics, and adds an Office licensing requirement. Recommend this **only** if the
formatting fidelity is a hard business requirement that cannot be met in code.

> **Decision required from the client/business:** *Is automatic preservation of the NOT-EQUAL
> conditional-formatting highlights a hard requirement?* — Yes → Recommendation B (Windows VM + Office).
> No / can be re-applied in code → Recommendation A (Container Apps). All resource lists below assume
> **Recommendation A** unless noted, and call out the VM delta.

---

## 3. Client provisioning checklist

### Group 1 — Microsoft Entra (identity & Graph permissions) — *target-state, required for the build*

- [ ] **Entra app registration** (single-tenant) for the service.
- [ ] **Microsoft Graph — application permission** for SharePoint access:
  - **Least-privilege (recommended): `Sites.Selected`**, then grant this app **write** access to the
    **one specific SharePoint site** (per-site grant via Graph). Scopes access to only that site.
  - **Simpler alternative: `Sites.ReadWrite.All`** (tenant-wide read/write to all sites) — easier to
    set up, broader blast radius; use only if `Sites.Selected` per-site administration is not workable.
  - The app needs **read** (source files) **and write** (write back the output copy), so grant
    read-write, not read-only. *(Inferred from the current script writing an output file; confirm the
    output destination — see §6.)*
- [ ] **Admin consent** granted for the application permission (app-only permissions require it).
- [ ] **Credential for the app**, in order of preference:
  1. **Managed identity + Graph app-role assignment** — *no secret to store* (preferred; works with
     Container Apps/App Service). 
  2. **Certificate** (uploaded to the app registration; private key in Key Vault).
  3. **Client secret** (stored in Key Vault; note expiry/rotation).

### Group 1b — Microsoft Entra (user sign-in to the UI) — *target-state, required because the brief adds a UI*

There are **two distinct identity concerns** — keep them separate:
- **App → SharePoint** = **app-only** Graph credential (Group 1 above). Unchanged by the UI.
- **User → UI** = **interactive user sign-in** so only authorized staff can open the UI and trigger a
  run. This is **new**, introduced by the UI requirement.

- [ ] **App registration (or built-in auth) for user sign-in**, e.g. platform authentication
  (Container Apps / App Service built-in Entra auth) protecting the UI and trigger endpoint.
  - Delegated sign-in (`openid` / `profile` / `User.Read`) — **only to identify and authorize the
    user**; it is **not** used to reach SharePoint (that stays app-only).
  - [ ] Restrict who can sign in (assigned users/group, or an app-role) so not everyone in the tenant
    can trigger runs. *(Inferred best practice; confirm the authorized-user set — §6.)*
- [ ] Decide whether user sign-in reuses the same app registration as the Graph app or a **separate**
  one. *(Recommended: separate — one app for user auth, one for app-only Graph — for clean
  least-privilege. Inferred.)*

### Group 2 — Compute (Recommendation A: Container Apps)

- [ ] **Azure Container Apps** app hosting the FastAPI service.
  - Ingress: HTTP; **internal** if callers are inside the tenant/VNet, **external** (public) only if an
    outside trigger needs it — protect either way (see §5 networking). *(Trigger source is an open
    question — §6.)*
  - **Scaling:** min replicas **0** (scale-to-zero) if an initial cold-start on trigger is acceptable,
    or min **1** to avoid cold start; scale rule on HTTP concurrency. Sequential Excel-style processing
    means **do not** rely on many parallel replicas for a single run.
  - **System-assigned managed identity** enabled (for Graph + Key Vault/ACR access).
- [ ] **Azure Container Apps Environment** (hosts the app; VNet-injected if private networking is required).
- [ ] **Azure Container Registry (ACR)** to store the application image.
- [ ] **Log Analytics workspace** (backs Container Apps logs + Application Insights).
- [ ] *(Optional, for heavy/long runs)* **Container Apps Job** for the processing run, with the HTTP app
  as the trigger/status surface.

> **Recommendation B delta (Windows VM path):** replace the four items above with a **Windows VM / VMSS**
> sized for Excel + **Microsoft Office licensing**, a process manager for the FastAPI app, and your
> standard VM diagnostics. ACR/Environment not needed.

### Group 3 — Supporting resources (only what the code genuinely needs)

- [ ] **Azure Key Vault** — **required only if** using certificate or client-secret auth (Group 1,
  option 2/3). **Not needed** if managed identity (option 1) is used and no other secrets exist.
- [ ] **Application Insights** — **recommended** for a production HTTP service (request tracing,
  failures, run metrics). *Flagged as a recommendation, not a code dependency.*
- [ ] **Storage account / database / cache — NOT required.** The code keeps **no persistent state**;
  it produces a single output workbook and holds nothing else. Do **not** provision a database or cache.
  *(Only consider a Storage account if you later need durable scratch space for very large files or an
  audit trail of outputs — not required by current code.)*

---

## 4. Networking requirements

- **Outbound (required, target-state):** HTTPS (443) to **`login.microsoftonline.com`** (token) and
  **`graph.microsoft.com`** (SharePoint file I/O). If egress is locked down, allow-list these.
- **Inbound:** the FastAPI service needs an ingress endpoint that **end users reach to load the UI and
  trigger runs** (the brief makes this human-facing, not machine-only).
  - **If UI users are all on the corporate network/VPN/VNet:** **internal ingress + VNet integration**,
    reachable privately — least exposure. Preferred where feasible.
  - **If UI users need to reach it over the internet:** external ingress is acceptable **only** behind
    **Entra sign-in** (Group 1b) — never an unauthenticated public endpoint. Consider fronting with
    Application Gateway/WAF or Front Door if public.
- **Private access to PaaS (optional hardening):** Private Endpoints for ACR/Key Vault if the client's
  policy requires no public data-plane access.
- **Where the UI users sit** (internal-only vs internet) decides internal vs external ingress — see §6.

---

## 5. Required configuration & environment variables (names only)

**Currently in code (CONFIG dict / CLI):**
- `FOLDER` · `TARGET_NAME_TOKEN` · `SOURCE_TAB` · `FORM_HEADER` · `LINE_HEADER` · `CHECK_HEADER` ·
  `EXCEL_VISIBLE`

**Target-state, to be added for the Azure/Graph build (not yet present in code — inferred):**
- Identity / Graph:
  - `AZURE_TENANT_ID`
  - `AZURE_CLIENT_ID`
  - `AZURE_CLIENT_SECRET` *(only if secret auth)* **or** `AZURE_CLIENT_CERTIFICATE_*` *(only if cert auth)* — **omit entirely if using managed identity**
- SharePoint targeting — **note:** the brief says the **user supplies these paths through the UI at
  run time**, so they are primarily **request inputs**, not static settings. Any env vars here are just
  **optional defaults / allow-list bounds**:
  - `SHAREPOINT_SITE_ID` (or `SHAREPOINT_HOSTNAME` + `SHAREPOINT_SITE_PATH`) — default/allowed site
  - `SHAREPOINT_DRIVE_ID` (or document-library name)
  - `SHAREPOINT_SOURCE_FOLDER_PATH`
  - `SHAREPOINT_OUTPUT_FOLDER_PATH` *(pending §6 confirmation of where output goes)*
  - *(If `Sites.Selected` is used, user-entered sites are constrained to whatever site(s) the app was
    granted — worth validating user input against that in the app.)*
- UI / user auth (target-state, from the UI requirement):
  - `AUTH_*` / built-in-auth settings for the Entra user sign-in (client id / allowed audience /
    allowed users or group) — exact names depend on the platform auth chosen (Group 1b).
- Service / ops:
  - `APPLICATIONINSIGHTS_CONNECTION_STRING` *(if App Insights used)*
  - `LOG_LEVEL`
  - `PORT` / ASGI bind settings

> **Never** commit real values; source secrets from Key Vault or managed identity, non-secret config
> from app settings.

---

## 6. Assumptions & open questions

1. **Excel-engine decision (blocking).** Does the red NOT-EQUAL **conditional formatting have to be
   preserved automatically** (→ xlwings → Windows VM + Office, Recommendation B), or can it be
   re-applied in code (→ openpyxl → Container Apps, Recommendation A)? Everything downstream depends
   on this.
2. **Scope of the build.** The brief's FastAPI + Graph + app-only auth is **not in the codebase**; this
   document scopes infra for building it. Confirm that new development (HTTP layer + Graph I/O) is in
   scope, not just "lift-and-shift the script."
3. **Trigger / UI.** The brief resolves this to a **user-initiated web UI**. Confirm: (a) is a
   **scheduled/automated** trigger *also* wanted alongside the UI, or is manual-only sufficient?
   (b) Is the UI served by the **same FastAPI container** (recommended, no extra host) or a
   **separate SPA** (e.g. Static Web Apps)? (c) **Who are the UI users**, are they **internal-only or
   internet-facing**, and which users/group should be **authorized** to sign in and trigger runs?
4. **Output destination.** Current code writes the timestamped copy **next to the source**. In the
   target, does output go back to the **same SharePoint library**, a different folder, or elsewhere?
   Drives the Graph **write** permission and folder config.
5. **Permission breadth.** Confirm `Sites.Selected` (per-site) is acceptable to the client's Entra
   admins; fall back to `Sites.ReadWrite.All` only if per-site grants can't be administered.
6. **Concurrency & volume.** Expected number of entity files and periods per run, and target run-time,
   to size the compute and decide sync-endpoint vs background-job pattern. *(Code processes sequentially;
   no volume assumptions are encoded.)*
7. **Python version.** Standardize the target (repo venv is 3.13; README says 3.10+).
