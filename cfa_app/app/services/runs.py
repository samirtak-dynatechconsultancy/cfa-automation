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
    load_master_pair,
    read_source_from_bytes,
    save_workbook_to_bytes,
    transfer_into_master,
    validate_master,
    verify_output,
)
from ..core.models import DetectionConfig, FileResult, RunResult
from ..core.naming import folder_period, folder_year
from ..graph.client import GraphClient
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
    """Build the output base name, e.g. 'CFA verification_P5_2026' or 'CFA verification_P4-P6_2026'.

    Strips any trailing period/year label already on the master name so it isn't duplicated.
    """
    base = params.master_name[:-5] if params.master_name.lower().endswith(".xlsx") \
        else params.master_name
    base = re.sub(r"[ _-]+P\d{1,2}([ _-]+\d{4})?\s*$", "", base, flags=re.IGNORECASE).strip(" _-")
    if params.period_from == params.period_to:
        period_part = f"P{params.period_from}"
    else:
        period_part = f"P{params.period_from}-P{params.period_to}"
    return f"{base}_{period_part}_{params.year}"


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

    def _unique_output_name(self, drive_id: str, folder_id: str, stem: str) -> str:
        """'{stem}.xlsx', or '{stem}_v1.xlsx', '{stem}_v2.xlsx', … if that name already exists."""
        try:
            existing = {it["name"].lower() for it in self._graph.list_children(drive_id, folder_id)
                        if not it["is_folder"]}
        except Exception:  # if listing fails, fall back to the plain name
            existing = set()
        candidate = f"{stem}.xlsx"
        if candidate.lower() not in existing:
            return candidate
        n = 1
        while f"{stem}_v{n}.xlsx".lower() in existing:
            n += 1
        return f"{stem}_v{n}.xlsx"

    # -- worker ------------------------------------------------------------
    def _execute(self, run_id: str, params: RunParams) -> None:
        result = self._runs[run_id]
        result.status = "running"
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
            # 1. Load master (a reusable, year-agnostic template).
            stage("Opening the master workbook…")
            master_bytes = self._graph.download_item(params.master_drive_id, params.master_item_id)
            values_wb, write_wb = load_master_pair(master_bytes)
            emit(f"master loaded ({len(master_bytes):,} bytes)")
            problems = validate_master(values_wb, params.year, params.period_from, params.master_name)
            if problems:
                result.status = "error"
                result.message = "Master check failed: " + "; ".join(problems)
                stage("Couldn't use the master workbook.")
                emit("ERROR: " + result.message)
                return

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
                    if src.month != period or src.year != params.year:
                        fr.messages.append(
                            f"file says P{src.month}/{src.year}, expected P{period}/{params.year}")
                        result.files.append(fr)
                        emit(f"  SKIP {sel.entity} {sel.currency} — period/year mismatch ({sel.name})")
                        continue
                    src.currency = sel.currency
                    fr = transfer_into_master(values_wb, write_wb, src, cfg)
                    fr.currency = sel.currency
                    fr.path = sel.folder_path
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

            # 4. Save + upload with a period/year name (versioned on collision).
            stage("Saving the filled workbook…")
            out_bytes = save_workbook_to_bytes(write_wb)
            stem = _output_stem(params)
            out_name = self._unique_output_name(
                params.output_drive_id, params.output_folder_id, stem)
            stage("Uploading the result to the output folder…")
            emit(f"uploading '{out_name}'")
            uploaded = self._graph.upload_to_folder(
                params.output_drive_id, params.output_folder_id, out_name, out_bytes)
            result.output_name = out_name
            result.output_url = uploaded.get("webUrl")
            emit(f"uploaded ({len(out_bytes):,} bytes)")

            # Independent verification of the saved copy.
            stage("Double-checking every value…")
            result.verify = verify_output(out_bytes, master_bytes, sources, cfg)
            emit(f"verified {result.verify.checked} cells — {result.verify.mismatches} mismatch(es)")

            done = sum(1 for f in result.files if f.status == "done")
            result.status = "done"
            result.processed = result.total
            mm = result.verify.mismatches
            result.message = (f"{done} entity file(s) transferred and "
                              f"{result.verify.checked:,} values checked — "
                              + ("all correct." if mm == 0 else f"{mm} mismatch(es) found."))
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
            result.finished_at = _dt.datetime.now().isoformat(timespec="seconds")
            try:
                self._run_log.log_run(result, params)
            except Exception:  # logging must never break a run
                log.exception("run %s: run-history logging raised", run_id)
