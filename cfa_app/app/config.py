"""Application configuration, loaded from environment / .env (no secrets in code)."""

from __future__ import annotations

import os
from functools import lru_cache

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:  # pragma: no cover - dotenv optional
    pass


class Settings:
    # --- Microsoft Graph (app-only, client credentials) ---
    tenant_id: str = os.getenv("GRAPH_TENANT_ID", "")
    client_id: str = os.getenv("GRAPH_CLIENT_ID", "")
    client_secret: str = os.getenv("GRAPH_CLIENT_SECRET", "")
    graph_base: str = os.getenv("GRAPH_BASE_URL", "https://graph.microsoft.com/v1.0")
    authority: str = os.getenv(
        "GRAPH_AUTHORITY",
        f"https://login.microsoftonline.com/{os.getenv('GRAPH_TENANT_ID', '')}",
    )

    # --- Settings store: a SharePoint List (auto-created) on this site ---
    # Provide the site URL only; the list is created if it doesn't exist.
    settings_site_url: str = os.getenv("SETTINGS_SITE_URL", "")   # e.g. https://host/sites/Finance
    settings_list_name: str = os.getenv("SETTINGS_LIST_NAME", "CFA App Settings")

    # --- Run history: a SharePoint List on the same site (created manually) ---
    runs_list_name: str = os.getenv("RUNS_LIST_NAME", "CFA App Runs")

    # --- Service ---
    app_host: str = os.getenv("APP_HOST", "127.0.0.1")
    app_port: int = int(os.getenv("APP_PORT", "8000"))
    log_level: str = os.getenv("LOG_LEVEL", "INFO")

    @property
    def graph_configured(self) -> bool:
        return bool(self.tenant_id and self.client_id and self.client_secret)


@lru_cache
def get_settings() -> Settings:
    return Settings()
