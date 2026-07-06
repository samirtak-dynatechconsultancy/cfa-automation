"""Run manager: orchestrates a transfer run as an async background job with status polling.

In-memory registry + a worker thread per run. Fine for the local single-user tool; when this
moves to central hosting it would be swapped for a queue/Container Apps Job (the RunManager
interface is what callers depend on, so that swap is contained here).
"""

from __future__ import annotations

import datetime as _dt
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


class RunManager:
    def __init__(self, graph: GraphClient, store: SettingsStore, run_log: RunLogStore):
        self._graph = graph
        self._store = store
        self._run_log = run_log
        self._runs: dict[str, RunResult] = {}
        self._lock = threading.Lock()

    def start(self, params: RunParams) -> str:
        run_id = uuid.uuid4().hex[:12]
        result = RunResult(run_id=run_id, status="queued")
        with self._lock:
            self._runs[run_id] = result
        log.info("run %s queued (master=%s, P%d-P%d)", run_id, params.master_name,
                 params.period_from, params.period_to)
        threading.Thread(target=self._execute, args=(run_id, params), daemon=True).start()
        return run_id

    def get(self, run_id: str) -> RunResult | None:
        with self._lock:
            return self._runs.get(run_id)

    # -- worker ------------------------------------------------------------
    def _execute(self, run_id: str, params: RunParams) -> None:
        result = self._runs[run_id]
        result.status = "running"
        result.started_at = _dt.datetime.now().isoformat(timespec="seconds")
        cfg: DetectionConfig = self._store.detection_config()
        periods = list(range(params.period_from, params.period_to + 1))

        def emit(msg: str) -> None:
            ts = _dt.datetime.now().strftime("%H:%M:%S")
            result.logs.append(f"{ts}  {msg}")
            log.info("run %s: %s", run_id, msg)

        emit(f"started — {params.year} P{params.period_from}-P{params.period_to}")
        try:
            # 1. Load master (a reusable, year-agnostic template).
            emit("downloading master workbook…")
            master_bytes = self._graph.download_item(params.master_drive_id, params.master_item_id)
            values_wb, write_wb = load_master_pair(master_bytes)
            emit(f"master loaded ({len(master_bytes):,} bytes)")
            problems = validate_master(values_wb, params.year, params.period_from, params.master_name)
            if problems:
                result.status = "error"
                result.message = "Master check failed: " + "; ".join(problems)
                emit("ERROR: " + result.message)
                return

            # 2. Scan the source tree ONCE, pruning branches for other years/periods so we only
            #    descend into the relevant folders instead of the whole tree.
            period_set = set(periods)

            def keep_folder(name: str) -> bool:
                fy = folder_year(name)
                if fy is not None and fy != params.year:
                    return False                       # a different year's folder
                fp = folder_period(name)
                if fp is not None and fp not in period_set:
                    return False                       # a period folder outside the range
                return True                            # region/entity/year-match/unknown -> descend

            emit(f"scanning for {params.year} P{params.period_from}-P{params.period_to} "
                 "(skipping other years/periods)…")
            last = {"n": 0}

            def scan_progress(state):
                if state["folders"] - last["n"] >= 25:
                    last["n"] = state["folders"]
                    emit(f"  scanning… {state['folders']} folders, {state['files']} files so far")

            all_files = self._graph.walk_files(params.source_drive_id, params.source_folder_id,
                                               on_progress=scan_progress, keep_folder=keep_folder)
            emit(f"scan complete: {len(all_files)} file(s) found")

            sources = []
            for period in periods:
                selections = select_sources(all_files, params.year, period)
                emit(f"P{period}: selected {len(selections)} entity file(s)")
                for sel in selections:
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

            if not result.files:
                result.status = "error"
                result.message = (f"No CFA source files found for {params.year} / "
                                  f"P{params.period_from}-P{params.period_to} under the selected folder.")
                emit("ERROR: " + result.message)
                return

            if not sources:
                result.status = "error"
                result.message = "No source produced any written lines; nothing to save."
                emit("ERROR: " + result.message)
                return

            # Save timestamped copy and upload to the output folder.
            emit("saving output workbook…")
            out_bytes = save_workbook_to_bytes(write_wb)
            ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            base = params.master_name[:-5] if params.master_name.lower().endswith(".xlsx") \
                else params.master_name
            out_name = f"{base} _autofilled_{ts}.xlsx"
            emit(f"uploading '{out_name}' to the output folder…")
            uploaded = self._graph.upload_to_folder(
                params.output_drive_id, params.output_folder_id, out_name, out_bytes)
            result.output_name = out_name
            result.output_url = uploaded.get("webUrl")
            emit(f"uploaded ({len(out_bytes):,} bytes)")

            # Independent verification of the saved copy.
            emit("verifying output…")
            result.verify = verify_output(out_bytes, master_bytes, sources, cfg)
            emit(f"verified {result.verify.checked} cells — {result.verify.mismatches} mismatch(es)")

            done = sum(1 for f in result.files if f.status == "done")
            result.status = "done"
            result.message = (f"{done} file(s) transferred, "
                              f"{result.verify.checked} cells verified, "
                              f"{result.verify.mismatches} mismatch(es).")
            emit("DONE — " + result.message)
        except Exception as e:
            result.status = "error"
            result.message = f"{type(e).__name__}: {e}"
            ts = _dt.datetime.now().strftime("%H:%M:%S")
            result.logs.append(f"{ts}  FAILED: {result.message}")
            log.exception("run %s FAILED", run_id)
        finally:
            result.finished_at = _dt.datetime.now().isoformat(timespec="seconds")
            try:
                self._run_log.log_run(result, params)
            except Exception:  # logging must never break a run
                log.exception("run %s: run-history logging raised", run_id)
