"""Run manager: orchestrates a transfer run as an async background job with status polling.

In-memory registry + a worker thread per run. Fine for the local single-user tool; when this
moves to central hosting it would be swapped for a queue/Container Apps Job (the RunManager
interface is what callers depend on, so that swap is contained here).
"""

from __future__ import annotations

import datetime as _dt
import re
import threading
import uuid
from dataclasses import dataclass

from ..core.discovery import select_sources
from ..core.engine import (
    hide_configured_rows,
    hide_period_gridlines,
    keep_only_periods,
    load_master_pair,
    load_write_workbook,
    read_source_from_bytes,
    save_workbook_to_bytes,
    transfer_into_master,
    validate_master,
    verify_output,
)
from ..core.models import DetectionConfig, FileResult, RunResult
from ..core.naming import folder_period, folder_region, folder_year
from ..graph.client import GraphClient, GraphError, GraphLockedError
from ..logging_config import get_logger
from .run_log import RunLogStore
from .settings_store import SettingsStore

log = get_logger("runs")


@dataclass
class RunParams:
    year: int
    period_from: int
    period_to: int
    source_drive_id: str
    source_folder_id: str          # recursion root (Year / Region / Company / Period folder)
    master_drive_id: str
    master_item_id: str
    master_name: str
    output_drive_id: str
    output_folder_id: str
    source_folder_name: str = ""
    output_folder_name: str = ""


def _output_stem(params: "RunParams") -> str:
    """Build the per-YEAR output base name, e.g. 'CFA verification_2026'.

    One file per year: every period run for that year appends its sheet into the same file, so the
    name carries only the year (any period/year label already on the master name is stripped first).
    """
    base = params.master_name[:-5] if params.master_name.lower().endswith(".xlsx") \
        else params.master_name
    base = re.sub(r"[ _-]+P\d{1,2}([ _-]+\d{4})?\s*$", "", base, flags=re.IGNORECASE).strip(" _-")
    return f"{base}_{params.year}"


class RunManager:
    def __init__(self, graph: GraphClient, store: SettingsStore, run_log: RunLogStore):
        self._graph = graph
        self._store = store
        self._run_log = run_log
        self._runs: dict[str, RunResult] = {}
        self._latest_id: str | None = None
        self._lock = threading.Lock()

    def start(self, params: RunParams) -> str:
        run_id = uuid.uuid4().hex[:12]
        result = RunResult(run_id=run_id, status="queued")
        with self._lock:
            self._runs[run_id] = result
            self._latest_id = run_id
        log.info("run %s queued (master=%s, P%d-P%d)", run_id, params.master_name,
                 params.period_from, params.period_to)
        threading.Thread(target=self._execute, args=(run_id, params), daemon=True).start()
        return run_id

    def get(self, run_id: str) -> RunResult | None:
        with self._lock:
            return self._runs.get(run_id)

    def latest(self) -> RunResult | None:
        """The most recently started run (used to re-attach after a refresh, incl. embedded)."""
        with self._lock:
            return self._runs.get(self._latest_id) if self._latest_id else None

    def _latest_output_file(self, drive_id: str, folder_id: str, stem: str) -> dict | None:
        """Most recently modified output for this year — the canonical '{stem}.xlsx' or a version
        '{stem}_v{n}.xlsx'. That file holds the newest set of periods, so appending onto it means a
        version created while the main file was locked is never lost from the next run.
        """
        pat = re.compile(rf"^{re.escape(stem)}(_v\d+)?\.xlsx$", re.IGNORECASE)
        best = None
        try:
            for it in self._graph.list_children(drive_id, folder_id):
                if it["is_folder"] or not pat.match(it["name"] or ""):
                    continue
                if best is None or (it.get("last_modified") or "") > (best.get("last_modified") or ""):
                    best = it
        except Exception:      # if listing fails, treat as not-present (a fresh file is created)
            return None
        return best

    def _versioned_output_name(self, drive_id: str, folder_id: str, stem: str) -> str:
        """First free '{stem}_v{n}.xlsx' name — used when the main year file is locked."""
        try:
            existing = {it["name"].lower() for it in self._graph.list_children(drive_id, folder_id)
                        if not it["is_folder"]}
        except Exception:
            existing = set()
        n = 1
        while f"{stem}_v{n}.xlsx".lower() in existing:
            n += 1
        return f"{stem}_v{n}.xlsx"

    # -- worker ------------------------------------------------------------
    def _execute(self, run_id: str, params: RunParams, *,
                 auto_version: bool = False, skip_probe: bool = False) -> None:
        result = self._runs[run_id]
        result.status = "running"
        if not result.started_at:
            result.started_at = _dt.datetime.now().isoformat(timespec="seconds")
        cfg: DetectionConfig = self._store.detection_config()
        periods = list(range(params.period_from, params.period_to + 1))

        def emit(msg: str) -> None:
            ts = _dt.datetime.now().strftime("%H:%M:%S")
            result.logs.append(f"{ts}  {msg}")     # technical detail (collapsible in the UI)
            log.info("run %s: %s", run_id, msg)

        def stage(msg: str) -> None:
            result.stage = msg                     # plain-language headline, replaced each time

        stage("Getting ready…")
        emit(f"started — {params.year} P{params.period_from}-P{params.period_to}")
        try:
            # 1. Load master (a reusable, year-agnostic template) — always the source of LABELS.
            stage("Opening the master workbook…")
            master_bytes = self._graph.download_item(params.master_drive_id, params.master_item_id)
            values_wb, template_write_wb = load_master_pair(master_bytes)
            emit(f"master loaded ({len(master_bytes):,} bytes)")
            problems = validate_master(values_wb, params.year, params.period_from, params.master_name)
            if problems:
                result.status = "error"
                result.message = "Master check failed: " + "; ".join(problems)
                stage("Couldn't use the master workbook.")
                emit("ERROR: " + result.message)
                return

            # One file per YEAR: if it already exists, append this run's period(s) into it;
            # otherwise start from the master template and strip the other periods afterwards.
            stem = _output_stem(params)
            out_name = stem + ".xlsx"
            existing = self._latest_output_file(
                params.output_drive_id, params.output_folder_id, stem)
            write_wb = template_write_wb
            initial = True
            if existing:
                try:
                    year_bytes = self._graph.download_item(
                        params.output_drive_id, existing["id"])
                except Exception as e:   # fall back to a fresh file from the template
                    emit(f"could not open '{existing['name']}' ({e}); creating fresh from template")
                    year_bytes = None
                if year_bytes is not None:
                    # UPFRONT lock check: test-write the year file BEFORE the long scan/transfer, so a
                    # file that's open in Excel is caught immediately (SharePoint can't report an Excel
                    # lock without a write). retries=0 = quick check. If locked, ask before any work.
                    if not skip_probe:
                        stage("Checking the output file is available…")
                        emit(f"checking whether '{out_name}' can be written")
                        try:
                            self._graph.upload_to_folder(
                                params.output_drive_id, params.output_folder_id,
                                out_name, year_bytes, retries=0)
                        except GraphLockedError:
                            result.pending = {"phase": "pre_run", "params": params,
                                              "out_name": out_name}
                            result.status = "awaiting"
                            result.message = (
                                f"'{out_name}' is open or locked — someone may have it open in Excel.")
                            result.stage = "The output file is locked — waiting for your choice."
                            emit("WAITING (before run): output file locked; asking whether to proceed")
                            return
                    write_wb = load_write_workbook(year_bytes)
                    initial = False
                    emit(f"appending onto latest output '{existing['name']}' "
                         f"({len(year_bytes):,} bytes)")

            # 2. Scan the source tree ONCE, pruning branches for other years/periods.
            period_set = set(periods)

            def keep_folder(name: str) -> bool:
                fy = folder_year(name)
                if fy is not None and fy != params.year:
                    return False
                fp = folder_period(name)
                if fp is not None and fp not in period_set:
                    return False
                return True

            stage("Looking through the folder for the entity files…")
            emit(f"scanning for {params.year} P{params.period_from}-P{params.period_to}")

            def scan_progress(state):
                stage(f"Looking through the folder… {state['files']} file(s) found so far")

            all_files = self._graph.walk_files(params.source_drive_id, params.source_folder_id,
                                               on_progress=scan_progress, keep_folder=keep_folder)
            emit(f"scan complete: {len(all_files)} file(s) found")

            # 3. Pick the file to use for each entity + currency, across the period range.
            stage("Choosing the latest file for each entity…")
            work = []   # list of (period, selection)
            for period in periods:
                for sel in select_sources(all_files, params.year, period):
                    work.append((period, sel))
            result.total = len(work)
            emit(f"selected {result.total} entity file(s) to process")

            if not work:
                result.status = "error"
                result.message = (f"No CFA source files found for {params.year} / "
                                  f"P{params.period_from}-P{params.period_to} under the selected folder.")
                stage("No matching files were found for the chosen year and period.")
                emit("ERROR: " + result.message)
                return

            sources = []
            for period, sel in work:
                result.processed += 1
                stage(f"Filling in entity {sel.entity} ({sel.currency}) — "
                      f"{result.processed} of {result.total}…")
                fr = FileResult(filename=sel.name, entity=sel.entity,
                                currency=sel.currency, path=sel.folder_path)
                try:
                    raw = self._graph.download_item(params.source_drive_id, sel.item_id)
                    src = read_source_from_bytes(raw, sel.name, cfg)
                    if src is None:
                        fr.messages.append("not a valid consistency-check source (no NUMBER/PERIOD/tab)")
                        result.files.append(fr)
                        emit(f"  SKIP {sel.entity} {sel.currency} — not a valid source ({sel.name})")
                        continue
                    # Entity comes from what's written in the sheet (the source's NUMBER cell) — that
                    # is the authoritative entity for the transfer. Period/year, however, stay pinned
                    # to the run's selected period so the values always land in the chosen period
                    # sheet (base_name = P{src.month}); a stale PERIOD cell must not misroute or drop
                    # a file. Currency likewise follows the filename (how the file was grouped).
                    file_period = f"P{src.month}/{src.year}"
                    period_mismatch = (src.month != period or src.year != params.year)
                    src.currency = sel.currency
                    src.month = period
                    src.year = params.year
                    src.source_url = sel.web_url          # hyperlinked on the entity number cell
                    region = folder_region(sel.folder_path, params.year)
                    fr = transfer_into_master(values_wb, write_wb, src, cfg, region=region)
                    fr.currency = sel.currency
                    fr.path = sel.folder_path
                    if period_mismatch:
                        fr.messages.append(
                            f"file's PERIOD cell says {file_period} but the run is for "
                            f"P{period}/{params.year} — used P{period}/{params.year}")
                        emit(f"  NOTE {fr.entity}: PERIOD cell says {file_period}, used P{period}")
                    result.files.append(fr)
                    if fr.status == "done":
                        sources.append(src)
                    if fr.entity_added:
                        notice = (f"Entity {fr.entity} ({sel.currency}) was not present in sheet "
                                  f"{fr.sheet} — a new column was added for it.")
                        result.notices.append(notice)
                        emit("NOTE: " + notice)
                    emit(f"  {fr.status.upper()} {fr.entity} {sel.currency} -> {fr.sheet} "
                         f"({fr.written} lines)  [{sel.full_path}]")
                except Exception as e:  # one bad file must not kill the batch
                    fr.status = "error"
                    fr.messages.append(f"error: {e!r}")
                    emit(f"  ERROR {sel.entity} {sel.currency}: {e}")
                    result.files.append(fr)
                    log.exception("run %s: ERROR on %s", run_id, sel.name)

            if not sources:
                result.status = "error"
                result.message = "No source produced any written lines; nothing to save."
                stage("Nothing could be written into the master.")
                emit("ERROR: " + result.message)
                return

            # 4. On a fresh year file, drop the template's other periods (keep only what we ran).
            if initial:
                removed = keep_only_periods(write_wb, period_set)
                if removed:
                    emit(f"removed template period sheet(s): {', '.join(removed)}")

            # 5. Hide the configured detail rows on every period sheet, then save + verify.
            hidden = hide_configured_rows(write_wb, cfg)
            if hidden:
                emit(f"hid {hidden} configured row(s) across the period sheet(s)")
            hide_period_gridlines(write_wb)    # clean boxless look, consistent across all periods
            stage("Saving the filled workbook…")
            # Repair the comment layer against whatever we loaded the workbook from (the existing
            # year file when appending, else the master template) so Excel Online accepts the file.
            comment_source = year_bytes if (not initial and year_bytes is not None) else master_bytes
            out_bytes = save_workbook_to_bytes(write_wb, source_bytes=comment_source)
            stage("Double-checking every value…")
            result.verify = verify_output(out_bytes, master_bytes, sources, cfg)
            emit(f"verified {result.verify.checked} cells — {result.verify.mismatches} mismatch(es)")

            done = sum(1 for f in result.files if f.status == "done")
            flagged = sum(f.highlighted for f in result.files)
            mm = result.verify.mismatches
            base_message = (f"{done} entity file(s) transferred and "
                            f"{result.verify.checked:,} values checked — "
                            + ("all correct." if mm == 0 else f"{mm} mismatch(es) found.")
                            + (f" {flagged} cell(s) flagged for large differences "
                               f"(|diff| > {cfg.diff_threshold:g})." if flagged else ""))
            result.processed = result.total

            # 6. Upload (overwrites the year file). If it's locked even after the retries, PAUSE and
            # ask the user whether to save this run as a new version instead of losing it.
            stage("Uploading the result to the output folder…")
            emit(f"uploading '{out_name}' ({'append' if not initial else 'new'})")
            target_name = out_name
            try:
                uploaded = self._graph.upload_to_folder(
                    params.output_drive_id, params.output_folder_id, out_name, out_bytes)
            except GraphLockedError:
                if auto_version:
                    # User already agreed up front to a new version — save it without re-asking.
                    target_name = self._versioned_output_name(
                        params.output_drive_id, params.output_folder_id, stem)
                    emit(f"'{out_name}' still locked — saving new version '{target_name}'")
                    stage("Saving a new version…")
                    uploaded = self._graph.upload_to_folder(
                        params.output_drive_id, params.output_folder_id, target_name, out_bytes)
                else:
                    # Locked between the pre-flight check and now (rare) — ask before versioning.
                    result.pending = {
                        "phase": "post_run", "bytes": out_bytes, "stem": stem, "out_name": out_name,
                        "drive_id": params.output_drive_id, "folder_id": params.output_folder_id,
                        "params": params, "base_message": base_message,
                    }
                    result.status = "awaiting"
                    result.message = (
                        f"'{out_name}' is open or locked — someone may have it open in Excel.")
                    result.stage = "The output file is locked — waiting for your choice."
                    emit("WAITING: output file locked; asking to close it and recheck")
                    return

            result.output_name = target_name
            result.output_url = uploaded.get("webUrl")
            result.status = "done"
            result.message = base_message
            if target_name != out_name:
                result.message += (f" '{out_name}' was locked, so this run was saved as "
                                   f"'{target_name}'.")
            stage("Done.")
            emit("DONE — " + result.message)
        except Exception as e:
            result.status = "error"
            result.message = f"{type(e).__name__}: {e}"
            result.stage = "Something went wrong."
            ts = _dt.datetime.now().strftime("%H:%M:%S")
            result.logs.append(f"{ts}  FAILED: {result.message}")
            log.exception("run %s FAILED", run_id)
        finally:
            if result.status != "awaiting":     # a paused run is finalised when the user resolves it
                self._finalize(result, params)

    def _finalize(self, result: RunResult, params: "RunParams") -> None:
        """Stamp the finish time and log the run to history (once, when it reaches a terminal state)."""
        result.finished_at = _dt.datetime.now().isoformat(timespec="seconds")
        try:
            self._run_log.log_run(result, params)
        except Exception:  # logging must never break a run
            log.exception("run %s: run-history logging raised", result.run_id)

    def resolve_locked(self, run_id: str, action: str) -> RunResult | None:
        """Answer the locked-output prompt for a paused run.

        action 'recheck' -> re-test the lock (e.g. after closing the file); if it's now free the run
                            continues into the canonical file, otherwise it stays paused.
        action 'version' -> save this run as a new version ('{stem}_v{n}.xlsx').
        action 'cancel'  -> nothing is written.
        """
        with self._lock:
            result = self._runs.get(run_id)
        if result is None or result.status != "awaiting" or not result.pending:
            return result
        p = result.pending
        phase = p.get("phase")

        if action == "cancel":
            result.pending = None
            result.status = "cancelled"
            result.message = (f"Not saved — '{p['out_name']}' is locked and you chose not to "
                              f"create a new version. Close the file and run again.")
            result.stage = "Cancelled — nothing was saved."
            self._finalize(result, p["params"])
            return result

        if action == "recheck":
            if phase == "pre_run":
                # Re-run from the top: the probe re-tests the lock. If free now, it goes to the
                # canonical file; if still locked, it pauses again with the same prompt.
                result.pending = None
                result.status = "running"
                result.stage = "Re-checking the output file…"
                threading.Thread(target=self._execute, args=(run_id, p["params"]),
                                 daemon=True).start()
                return result
            # post_run: the bytes are ready — just retry writing the canonical file.
            result.stage = "Re-checking the output file…"
            try:
                uploaded = self._graph.upload_to_folder(
                    p["drive_id"], p["folder_id"], p["out_name"], p["bytes"], retries=0)
            except GraphLockedError:
                result.stage = "Still open/locked — waiting for your choice."
                return result                              # keep the pending state + prompt
            result.pending = None
            result.output_name = p["out_name"]
            result.output_url = uploaded.get("webUrl")
            result.status = "done"
            result.message = p["base_message"]
            result.stage = "Done."
            self._finalize(result, p["params"])
            return result

        # action == 'version'
        result.pending = None
        if phase == "pre_run":
            # Agreed to a version before the work: run it now, saving a version if still locked.
            result.status = "running"
            result.stage = "Getting ready…"
            threading.Thread(target=self._execute, args=(run_id, p["params"]),
                             kwargs={"auto_version": True, "skip_probe": True}, daemon=True).start()
            return result
        # post_run version: the bytes are already produced — save them as a new version now.
        target = self._versioned_output_name(p["drive_id"], p["folder_id"], p["stem"])
        result.stage = "Saving a new version…"
        try:
            uploaded = self._graph.upload_to_folder(
                p["drive_id"], p["folder_id"], target, p["bytes"])
        except GraphError as e:
            result.status = "error"
            result.message = f"Could not save a new version: {e}"
            result.stage = "Something went wrong saving the new version."
            self._finalize(result, p["params"])
            return result
        result.output_name = target
        result.output_url = uploaded.get("webUrl")
        result.status = "done"
        result.message = (p["base_message"] + f" '{p['out_name']}' was locked, so this run was "
                          f"saved as a new version '{target}'.")
        result.stage = "Done."
        self._finalize(result, p["params"])
        return result
