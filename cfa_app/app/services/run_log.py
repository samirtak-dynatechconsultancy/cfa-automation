"""Append one row per run to a SharePoint List for audit/history.

The list is created MANUALLY in SharePoint (so no Sites.Manage.All is needed) — this writer only
finds it and adds items. If the list is missing or Graph isn't configured, run logging is skipped
with a warning and never blocks the actual transfer run.
"""

from __future__ import annotations

import threading

from ..config import Settings
from ..graph.client import GraphClient, GraphError
from ..logging_config import get_logger

log = get_logger("runlog")


class RunLogStore:
    def __init__(self, settings: Settings, graph: GraphClient):
        self._settings = settings
        self._graph = graph
        self._lock = threading.Lock()
        self._site_id: str | None = None
        self._list_id: str | None = None
        self._columns: set[str] = set()   # internal column names actually present on the list
        self._checked_missing = False   # avoid repeated lookups once known missing

    @property
    def _persisted(self) -> bool:
        return self._settings.graph_configured and bool(self._settings.settings_site_url)

    def _resolve_list(self) -> bool:
        if self._list_id:
            return True
        if not self._persisted or self._checked_missing:
            return False
        try:
            self._site_id = self._graph.resolve_site_by_url(self._settings.settings_site_url)
            list_id = self._graph.find_list(self._site_id, self._settings.runs_list_name)
            if not list_id:
                log.warning("run log list '%s' not found on site — create it to enable run history",
                            self._settings.runs_list_name)
                self._checked_missing = True
                return False
            self._list_id = list_id
            try:
                self._columns = set(self._graph.list_columns(self._site_id, list_id))
            except GraphError:
                self._columns = set()
            log.info("run log: using list '%s' (%d columns)",
                     self._settings.runs_list_name, len(self._columns))
            return True
        except GraphError:
            log.warning("run log: could not resolve list", exc_info=True)
            return False

    def log_run(self, result, params) -> None:
        """Best-effort: write one row for a finished run. Never raises."""
        with self._lock:
            if not self._resolve_list():
                return
            files_total = len(result.files)
            files_done = sum(1 for f in result.files if f.status == "done")
            files_skipped = sum(1 for f in result.files if f.status == "skipped")
            files_error = sum(1 for f in result.files if f.status == "error")

            errors = []
            if result.status == "error" and result.message:
                errors.append(result.message)
            for f in result.files:
                if f.status in ("skipped", "error") and f.messages:
                    errors.append(f"{f.filename}: {f.messages[-1]}")

            fields = {
                "Title": result.run_id,
                "Status": result.status,
                "Year": getattr(params, "year", "") or "",
                "Period": getattr(params, "period_from", "") or "",
                "PeriodTo": getattr(params, "period_to", "") or "",
                "StartedAt": result.started_at or "",
                "FinishedAt": result.finished_at or "",
                "MasterName": params.master_name or "",
                "SourceFolder": getattr(params, "source_folder_name", "") or "",
                "OutputName": result.output_name or "",
                "OutputUrl": result.output_url or "",
                "FilesTotal": files_total,
                "FilesDone": files_done,
                "FilesSkipped": files_skipped,
                "FilesError": files_error,
                "CellsVerified": result.verify.checked if result.verify else 0,
                "Mismatches": result.verify.mismatches if result.verify else 0,
                "Message": result.message or "",
                "Errors": "; ".join(errors)[:4000],
            }
            # Only send fields that exist as columns on the list (Title always kept), so the writer
            # adapts to whatever columns you created instead of failing on unknown fields.
            if self._columns:
                dropped = [k for k in fields if k != "Title" and k not in self._columns]
                fields = {k: v for k, v in fields.items()
                          if k == "Title" or k in self._columns}
                if dropped:
                    log.info("run log: skipping columns not on the list: %s", ", ".join(dropped))
            try:
                self._graph.create_list_item_fields(self._site_id, self._list_id, fields)
                log.info("run %s logged to SharePoint history", result.run_id)
            except GraphError:
                log.warning("run %s: could not write to run history", result.run_id, exc_info=True)
