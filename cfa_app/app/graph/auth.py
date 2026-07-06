"""App-only (client credentials) token acquisition for Microsoft Graph, with simple caching."""

from __future__ import annotations

import threading
import time

import msal

from ..config import Settings
from ..logging_config import get_logger

log = get_logger("graph.auth")

_SCOPE = ["https://graph.microsoft.com/.default"]


class TokenProvider:
    def __init__(self, settings: Settings):
        self._settings = settings
        self._lock = threading.Lock()
        self._app: msal.ConfidentialClientApplication | None = None
        self._token: str | None = None
        self._expires_at: float = 0.0

    def _ensure_app(self) -> msal.ConfidentialClientApplication:
        if self._app is None:
            if not self._settings.graph_configured:
                raise RuntimeError(
                    "Graph is not configured. Set GRAPH_TENANT_ID, GRAPH_CLIENT_ID and "
                    "GRAPH_CLIENT_SECRET in the .env file."
                )
            self._app = msal.ConfidentialClientApplication(
                client_id=self._settings.client_id,
                authority=self._settings.authority,
                client_credential=self._settings.client_secret,
            )
        return self._app

    def get_token(self) -> str:
        with self._lock:
            if self._token and time.time() < self._expires_at - 60:
                return self._token
            app = self._ensure_app()
            result = app.acquire_token_for_client(scopes=_SCOPE)
            if "access_token" not in result:
                log.error("Graph token request failed: %s: %s",
                          result.get("error"), result.get("error_description"))
                raise RuntimeError(
                    "Failed to acquire Graph token: "
                    f"{result.get('error')}: {result.get('error_description')}"
                )
            self._token = result["access_token"]
            self._expires_at = time.time() + int(result.get("expires_in", 3600))
            log.info("acquired Graph token (valid ~%ss)", int(result.get("expires_in", 3600)))
            return self._token
