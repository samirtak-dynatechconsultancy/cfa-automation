"""Persist admin defaults + detection config in a SharePoint List (key/value items).

Only a settings SITE URL is required (SETTINGS_SITE_URL). The list is created automatically on
first use if it doesn't exist. Falls back to an in-memory copy when Graph / the site URL is not
configured, so the UI still works for a demo.
"""

from __future__ import annotations

import threading

from ..config import Settings
from ..core.models import DetectionConfig
from ..graph.client import GraphClient, GraphError
from ..logging_config import get_logger

log = get_logger("settings")

DEFAULT_KEYS = [
    "default_year", "default_period",
    "source_link", "source_drive_id", "source_folder_id", "source_folder_name",
    "master_link", "master_drive_id", "master_item_id", "master_name",
    "output_link", "output_drive_id", "output_folder_id", "output_folder_name",
]
DETECTION_KEYS = list(DetectionConfig().to_mapping().keys())


class SettingsStore:
    def __init__(self, settings: Settings, graph: GraphClient):
        self._settings = settings
        self._graph = graph
        self._lock = threading.Lock()
        self._cache: dict | None = None
        self._site_id: str | None = None
        self._list_id: str | None = None
        self._item_ids: dict[str, str] = {}     # key -> list item id

    # -- persistence location ---------------------------------------------
    @property
    def _location_configured(self) -> bool:
        return bool(self._settings.settings_site_url)

    @property
    def persisted(self) -> bool:
        return self._settings.graph_configured and self._location_configured

    def _ensure_list(self) -> None:
        """Resolve the site and the (auto-created) settings list. Cached after first call."""
        if self._list_id:
            return
        self._site_id = self._graph.resolve_site_by_url(self._settings.settings_site_url)
        existing = self._graph.find_list(self._site_id, self._settings.settings_list_name)
        if existing:
            self._list_id = existing
            log.info("settings: using existing list '%s' on site %s",
                     self._settings.settings_list_name, self._site_id)
        else:
            self._list_id = self._graph.create_list(
                self._site_id, self._settings.settings_list_name)
            log.info("settings: created list '%s' on site %s",
                     self._settings.settings_list_name, self._site_id)

    # -- load / save -------------------------------------------------------
    def load(self) -> dict:
        with self._lock:
            if self._cache is not None:
                return dict(self._cache)
            data = self._defaults()
            if self.persisted:
                try:
                    self._ensure_list()
                    self._item_ids = {}
                    for item in self._graph.list_items(self._site_id, self._list_id):
                        if item["key"]:
                            self._item_ids[item["key"]] = item["id"]
                            data[item["key"]] = item["value"] or ""
                    log.info("settings: loaded %d key(s) from SharePoint list", len(self._item_ids))
                except GraphError:
                    log.warning("settings: could not load from SharePoint; using defaults",
                                exc_info=True)
            self._cache = data
            return dict(data)

    def save(self, incoming: dict) -> dict:
        with self._lock:
            data = self._defaults()
            data.update({k: v for k, v in incoming.items() if k in DEFAULT_KEYS + DETECTION_KEYS})
            if self.persisted:
                self._ensure_list()
                if not self._item_ids:
                    for item in self._graph.list_items(self._site_id, self._list_id):
                        if item["key"]:
                            self._item_ids[item["key"]] = item["id"]
                for key in DEFAULT_KEYS + DETECTION_KEYS:
                    value = str(data.get(key, ""))
                    if key in self._item_ids:
                        self._graph.update_list_item(
                            self._site_id, self._list_id, self._item_ids[key], value)
                    else:
                        created = self._graph.add_list_item(
                            self._site_id, self._list_id, key, value)
                        self._item_ids[key] = created["id"]
                log.info("settings: saved %d key(s) to SharePoint list",
                         len(DEFAULT_KEYS + DETECTION_KEYS))
            else:
                log.info("settings: saved in memory only (not persisted)")
            self._cache = data
            return dict(data)

    # -- helpers -----------------------------------------------------------
    def detection_config(self) -> DetectionConfig:
        return DetectionConfig.from_mapping(self.load())

    def _defaults(self) -> dict:
        data = {k: "" for k in DEFAULT_KEYS}
        data.update({k: str(v) for k, v in DetectionConfig().to_mapping().items()})
        return data
