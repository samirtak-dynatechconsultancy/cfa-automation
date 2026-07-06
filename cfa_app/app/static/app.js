"use strict";

// ---------------------------------------------------------------------------
// SharePoint browser
// ---------------------------------------------------------------------------
const browserState = {
  mode: "folder",       // "folder" | "file"
  target: "source",     // "source" | "master" | "output"
  siteId: null,
  driveId: null,
  stack: [],            // [{id, name}] breadcrumb of folders (root = [])
};

async function api(url, opts) {
  const r = await fetch(url, opts);
  if (!r.ok) {
    let detail = r.statusText;
    try { detail = (await r.json()).detail || detail; } catch (e) {}
    throw new Error(detail);
  }
  return r.json();
}

function el(id) { return document.getElementById(id); }

function openBrowser(mode, target) {
  browserState.mode = mode;
  browserState.target = target;
  browserState.siteId = null;
  browserState.driveId = null;
  browserState.stack = [];
  el("browserTitle").textContent =
    "Select a " + (mode === "file" ? "file (master workbook)" : "folder") + " — " + target;
  el("driveStep").classList.add("hidden");
  el("itemStep").classList.add("hidden");
  el("siteList").innerHTML = "";
  el("siteQuery").value = "";
  el("browser").classList.remove("hidden");
  searchSites();
}

function closeBrowser() { el("browser").classList.add("hidden"); }

let siteTimer = null;
function debouncedSiteSearch() {
  clearTimeout(siteTimer);
  siteTimer = setTimeout(searchSites, 350);
}

async function searchSites() {
  const q = el("siteQuery").value;
  const list = el("siteList");
  list.innerHTML = "<li class='muted'>Searching…</li>";
  try {
    const { sites } = await api("/api/sites?q=" + encodeURIComponent(q));
    list.innerHTML = "";
    if (!sites.length) { list.innerHTML = "<li class='muted'>No sites found.</li>"; return; }
    sites.forEach(s => {
      const li = document.createElement("li");
      li.textContent = s.name || s.webUrl;
      li.onclick = () => selectSite(s);
      list.appendChild(li);
    });
  } catch (e) { list.innerHTML = "<li class='err'>" + e.message + "</li>"; }
}

async function selectSite(site) {
  browserState.siteId = site.id;
  browserState.siteName = site.name;
  el("driveStep").classList.remove("hidden");
  el("itemStep").classList.add("hidden");
  const list = el("driveList");
  list.innerHTML = "<li class='muted'>Loading libraries…</li>";
  try {
    const { drives } = await api("/api/drives?site_id=" + encodeURIComponent(site.id));
    list.innerHTML = "";
    drives.forEach(d => {
      const li = document.createElement("li");
      li.textContent = d.name + (d.driveType ? "  (" + d.driveType + ")" : "");
      li.onclick = () => selectDrive(d);
      list.appendChild(li);
    });
  } catch (e) { list.innerHTML = "<li class='err'>" + e.message + "</li>"; }
}

function selectDrive(drive) {
  browserState.driveId = drive.id;
  browserState.stack = [];
  el("itemStep").classList.remove("hidden");
  el("folderPickBar").classList.toggle("hidden", browserState.mode !== "folder");
  loadItems();
}

function currentFolderId() {
  const s = browserState.stack;
  return s.length ? s[s.length - 1].id : null;
}

async function loadItems() {
  renderBreadcrumb();
  const list = el("itemList");
  list.innerHTML = "<li class='muted'>Loading…</li>";
  try {
    let url = "/api/items?drive_id=" + encodeURIComponent(browserState.driveId);
    const fid = currentFolderId();
    if (fid) url += "&item_id=" + encodeURIComponent(fid);
    const { items } = await api(url);
    list.innerHTML = "";
    const folders = items.filter(i => i.is_folder);
    const files = items.filter(i => !i.is_folder);
    folders.forEach(it => list.appendChild(itemRow(it, true)));
    if (browserState.mode === "file") files.forEach(it => list.appendChild(itemRow(it, false)));
    if (!list.children.length) list.innerHTML = "<li class='muted'>Empty.</li>";
  } catch (e) { list.innerHTML = "<li class='err'>" + e.message + "</li>"; }
}

function itemRow(it, isFolder) {
  const li = document.createElement("li");
  li.textContent = (isFolder ? "📁 " : "📄 ") + it.name;
  if (isFolder) {
    li.onclick = () => { browserState.stack.push({ id: it.id, name: it.name }); loadItems(); };
  } else {
    li.classList.add("file");
    li.onclick = () => pickFile(it);
  }
  return li;
}

function renderBreadcrumb() {
  const bc = el("breadcrumb");
  bc.innerHTML = "";
  const root = document.createElement("span");
  root.textContent = (browserState.siteName || "site") + " / root";
  root.className = "crumb";
  root.onclick = () => { browserState.stack = []; loadItems(); };
  bc.appendChild(root);
  browserState.stack.forEach((f, idx) => {
    const sep = document.createElement("span"); sep.textContent = " / "; bc.appendChild(sep);
    const c = document.createElement("span");
    c.textContent = f.name; c.className = "crumb";
    c.onclick = () => { browserState.stack = browserState.stack.slice(0, idx + 1); loadItems(); };
    bc.appendChild(c);
  });
}

function pickCurrentFolder() {
  const fid = currentFolderId();
  if (!fid) { alert("Navigate into a folder first (cannot select the library root)."); return; }
  const name = browserState.stack[browserState.stack.length - 1].name;
  setTarget(browserState.target, browserState.driveId, fid, name);
  closeBrowser();
}

function pickFile(it) {
  setTarget(browserState.target, browserState.driveId, it.id, it.name);
  closeBrowser();
}

function setTarget(target, driveId, itemId, name) {
  if (target === "source") {
    el("source_drive_id").value = driveId;
    el("source_folder_id").value = itemId;
    el("source_folder_name").textContent = name;
  } else if (target === "master") {
    el("master_drive_id").value = driveId;
    el("master_item_id").value = itemId;
    el("master_name").textContent = name;
  } else if (target === "output") {
    el("output_drive_id").value = driveId;
    el("output_folder_id").value = itemId;
    el("output_folder_name").textContent = name;
  }
}

// ---------------------------------------------------------------------------
// Run + polling
// ---------------------------------------------------------------------------
// Default the "to" period to "from" if the user hasn't set it (single-period convenience).
function syncPeriodTo() {
  const to = el("period_to");
  if (!to.value) to.value = el("period_from").value;
}

async function startRun() {
  const year = parseInt(el("year").value, 10);
  const pFrom = parseInt(el("period_from").value, 10);
  const pTo = parseInt(el("period_to").value || el("period_from").value, 10);
  if (!year || !pFrom) { alert("Please enter the Year and choose a Period (from)."); return; }
  if (pTo < pFrom) { alert("'Period to' must not be before 'Period from'."); return; }
  const body = {
    year: year,
    period_from: pFrom,
    period_to: pTo,
    source_drive_id: el("source_drive_id").value,
    source_folder_id: el("source_folder_id").value,
    master_drive_id: el("master_drive_id").value,
    master_item_id: el("master_item_id").value,
    master_name: el("master_name").textContent.trim(),
    output_drive_id: el("output_drive_id").value,
    output_folder_id: el("output_folder_id").value,
    source_folder_name: el("source_folder_name").textContent.trim(),
    output_folder_name: el("output_folder_name").textContent.trim(),
  };
  if (!body.source_folder_id || !body.master_item_id || !body.output_folder_id) {
    alert("Please select a source folder, a master workbook, and an output folder.");
    return;
  }
  el("runBtn").disabled = true;
  el("result").classList.remove("hidden");
  el("spinner").className = "spinner spin";
  el("spinner").textContent = "";
  el("stageText").textContent = "Starting…";
  el("progressWrap").classList.add("hidden");
  el("progressBar").style.width = "0%";
  el("runMessage").textContent = "";
  el("verifyBox").innerHTML = "";
  el("noticeBox").innerHTML = "";
  el("runLog").textContent = "";
  el("filesTable").innerHTML = "";
  try {
    const { run_id } = await api("/api/runs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    pollRun(run_id);
  } catch (e) {
    el("spinner").className = "spinner bad";
    el("spinner").textContent = "✕";
    el("stageText").textContent = "Something went wrong.";
    el("runMessage").textContent = e.message;
    el("runBtn").disabled = false;
  }
}

async function pollRun(runId) {
  try {
    const r = await api("/api/runs/" + runId);
    renderRun(r);
    if (r.status === "queued" || r.status === "running") {
      setTimeout(() => pollRun(runId), 1500);
    } else {
      el("runBtn").disabled = false;
    }
  } catch (e) {
    el("spinner").className = "spinner bad";
    el("spinner").textContent = "✕";
    el("stageText").textContent = "Something went wrong.";
    el("runMessage").textContent = e.message;
    el("runBtn").disabled = false;
  }
}

function renderRun(r) {
  const running = (r.status === "queued" || r.status === "running");

  // Plain-language headline + spinner/check/cross.
  el("stageText").textContent = r.stage || (running ? "Working…" : "");
  const sp = el("spinner");
  if (running) { sp.className = "spinner spin"; sp.textContent = ""; }
  else if (r.status === "done") { sp.className = "spinner ok"; sp.textContent = "✓"; }
  else { sp.className = "spinner bad"; sp.textContent = "✕"; }

  // Progress bar (X of Y) once the total is known.
  const pw = el("progressWrap");
  if (r.total > 0) {
    pw.classList.remove("hidden");
    const pct = Math.max(0, Math.min(100, Math.round((r.processed / r.total) * 100)));
    el("progressBar").style.width = pct + "%";
  } else {
    pw.classList.add("hidden");
  }

  // One-line summary once finished.
  el("runMessage").textContent = running ? "" : (r.message || "");

  // Notices (e.g. entity columns that had to be added).
  const noticeEl = el("noticeBox");
  if (noticeEl) {
    if (r.notices && r.notices.length) {
      noticeEl.innerHTML = "<div class='notices'><strong>Heads up</strong><ul>" +
        r.notices.map(n => "<li>" + n + "</li>").join("") + "</ul></div>";
    } else {
      noticeEl.innerHTML = "";
    }
  }

  // Live progress log — updates on every poll while the run is in flight.
  const logEl = el("runLog");
  if (logEl && r.logs) {
    const atBottom = logEl.scrollTop + logEl.clientHeight >= logEl.scrollHeight - 20;
    logEl.textContent = r.logs.join("\n");
    if (atBottom) logEl.scrollTop = logEl.scrollHeight;   // keep following the tail
  }

  const vb = el("verifyBox");
  if (r.verify) {
    const cls = r.verify.passed ? "pass" : "fail";
    let html = "<div class='verify " + cls + "'>Verification: " +
      (r.verify.passed ? "ALL PASS" : "FAILURES") +
      " — " + r.verify.checked + " cells checked, " + r.verify.mismatches + " mismatch(es)</div>";
    if (r.verify.samples && r.verify.samples.length) {
      html += "<ul class='samples'>" +
        r.verify.samples.map(s => "<li>" + s + "</li>").join("") + "</ul>";
    }
    vb.innerHTML = html;
  } else { vb.innerHTML = ""; }

  if (r.output_name) {
    const link = r.output_url
      ? "<a href='" + r.output_url + "' target='_blank' rel='noopener'>" + r.output_name + "</a>"
      : r.output_name;
    vb.innerHTML += "<div class='output'>Output: " + link + "</div>";
  }

  const t = el("filesTable");
  if (r.files && r.files.length) {
    let rows = "<tr><th>File</th><th>Folder (source path)</th><th>Entity</th><th>Cur</th>" +
               "<th>Sheet</th><th>Status</th><th>Written</th><th>Skipped</th><th>Notes</th></tr>";
    r.files.forEach(f => {
      const notes = (f.messages || []).slice(0, 3).join("; ");
      rows += "<tr class='" + f.status + "'><td>" + f.filename +
        "</td><td class='path'>" + (f.path || "") + "</td><td>" + (f.entity || "") +
        "</td><td>" + (f.currency || "") + "</td><td>" + (f.sheet || "") +
        "</td><td>" + f.status + "</td><td>" + f.written +
        "</td><td>" + f.skipped_lines + "</td><td class='notes'>" + notes + "</td></tr>";
    });
    t.innerHTML = rows;
  }
}

// ---------------------------------------------------------------------------
// SharePoint "Copy link" resolution (shared by run + admin pages)
// ---------------------------------------------------------------------------
const TARGET_FIELDS = {
  source: { drive: "source_drive_id", item: "source_folder_id", name: "source_folder_name" },
  master: { drive: "master_drive_id", item: "master_item_id", name: "master_name" },
  output: { drive: "output_drive_id", item: "output_folder_id", name: "output_folder_name" },
};

async function resolveLink(url, expect) {
  return api("/api/resolve", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ url: url, expect: expect }),
  });
}

function showResolved(target, ok, text) {
  const st = el("resolved_" + target);
  if (!st) return;
  st.textContent = (ok ? "✓ " : "✗ ") + text;
  st.className = "resolved " + (ok ? "ok" : "bad");
}

// Run page: paste link -> auto-resolve -> set the hidden ids + chosen name span.
async function resolveInline(target) {
  const input = el("link_" + target);
  if (!input.value.trim()) { showResolved(target, true, ""); return; }
  showResolved(target, true, "Resolving…");
  try {
    const r = await resolveLink(input.value, input.dataset.expect);
    setTarget(target, r.drive_id, r.item_id, r.name);
    showResolved(target, true, r.name + (r.is_folder ? " (folder)" : " (file)"));
  } catch (e) { showResolved(target, false, e.message); }
}

// Debounced auto-resolve as the user pastes/types a link (no button needed).
function debouncedInline(inputEl) {
  clearTimeout(inputEl._t);
  inputEl._t = setTimeout(() => resolveInline(inputEl.dataset.target), 500);
}

// Admin page: paste link -> fill the hidden form fields (by name).
async function adminResolve(inputEl) {
  const target = inputEl.dataset.target;
  const form = el("adminForm");
  if (!inputEl.value.trim()) { showResolved(target, true, ""); return; }
  showResolved(target, true, "Resolving…");
  try {
    const r = await resolveLink(inputEl.value, inputEl.dataset.expect);
    const f = TARGET_FIELDS[target];
    form.elements[f.drive].value = r.drive_id;
    form.elements[f.item].value = r.item_id;
    form.elements[f.name].value = r.name;
    showResolved(target, true, r.name + (r.is_folder ? " (folder)" : " (file)"));
  } catch (e) { showResolved(target, false, e.message); }
}

function debouncedResolve(inputEl) {
  clearTimeout(inputEl._t);
  inputEl._t = setTimeout(() => adminResolve(inputEl), 500);
}

// ---------------------------------------------------------------------------
// Admin
// ---------------------------------------------------------------------------
function hydrateAdmin() {
  const raw = el("values-json");
  if (!raw) return;
  let values = {};
  try { values = JSON.parse(raw.textContent); } catch (e) {}
  const form = el("adminForm");
  Object.entries(values).forEach(([k, v]) => {
    const field = form.elements[k];
    if (field) field.value = v;
  });
  // Reflect any previously-resolved targets without re-hitting Graph.
  ["source", "master", "output"].forEach(t => {
    const nm = form.elements[TARGET_FIELDS[t].name].value;
    if (nm) showResolved(t, true, nm);
  });
}

async function saveSettings(ev) {
  ev.preventDefault();
  const form = el("adminForm");
  const payload = {};
  Array.from(form.elements).forEach(f => { if (f.name) payload[f.name] = f.value; });
  const msg = el("saveMsg");
  msg.textContent = "Saving…";
  try {
    const res = await api("/api/settings", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    msg.textContent = res.persisted ? "Saved to SharePoint." : "Saved (in memory only).";
  } catch (e) { msg.textContent = "Error: " + e.message; }
  return false;
}
