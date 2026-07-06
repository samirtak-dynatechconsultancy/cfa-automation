"""Thin Microsoft Graph client for SharePoint browsing and file I/O (app-only)."""

from __future__ import annotations

import base64
from urllib.parse import unquote, urlparse

import httpx

from ..config import Settings
from ..logging_config import get_logger
from .auth import TokenProvider

log = get_logger("graph")

_UPLOAD_SIMPLE_LIMIT = 4 * 1024 * 1024   # 4 MiB: below this, PUT ...:/content works directly


class GraphError(RuntimeError):
    pass


class GraphClient:
    def __init__(self, settings: Settings, token_provider: TokenProvider):
        self._settings = settings
        self._tokens = token_provider
        self._base = settings.graph_base.rstrip("/")

    # -- low level ---------------------------------------------------------
    def _headers(self, extra: dict | None = None) -> dict:
        h = {"Authorization": f"Bearer {self._tokens.get_token()}"}
        if extra:
            h.update(extra)
        return h

    def _get(self, path: str, params: dict | None = None) -> dict:
        url = path if path.startswith("http") else f"{self._base}{path}"
        with httpx.Client(timeout=60) as c:
            r = c.get(url, headers=self._headers(), params=params)
        log.debug("GET %s -> %s", path, r.status_code)
        if r.status_code >= 400:
            raise GraphError(f"GET {url} -> {r.status_code}: {r.text[:500]}")
        return r.json()

    def _post(self, path: str, json: dict) -> dict:
        url = f"{self._base}{path}"
        with httpx.Client(timeout=60) as c:
            r = c.post(url, headers=self._headers({"Content-Type": "application/json"}), json=json)
        log.debug("POST %s -> %s", path, r.status_code)
        if r.status_code >= 400:
            raise GraphError(f"POST {url} -> {r.status_code}: {r.text[:500]}")
        return r.json()

    def _patch(self, path: str, json: dict) -> dict:
        url = f"{self._base}{path}"
        with httpx.Client(timeout=60) as c:
            r = c.patch(url, headers=self._headers({"Content-Type": "application/json"}), json=json)
        log.debug("PATCH %s -> %s", path, r.status_code)
        if r.status_code >= 400:
            raise GraphError(f"PATCH {url} -> {r.status_code}: {r.text[:500]}")
        return r.json()

    # -- sites / drives / items -------------------------------------------
    def search_sites(self, query: str) -> list[dict]:
        """Search sites across the tenant. Empty query returns a small default set."""
        q = query.strip() or "*"
        data = self._get("/sites", params={"search": q})
        out = []
        for s in data.get("value", []):
            out.append({
                "id": s.get("id"),
                "name": s.get("displayName") or s.get("name"),
                "webUrl": s.get("webUrl"),
            })
        return out

    def resolve_share_link(self, url: str) -> dict:
        """Resolve a SharePoint/OneDrive 'Copy link' URL into a driveItem (drive id + item id).

        Uses the Graph /shares endpoint with the URL encoded as a share id.
        """
        token = base64.urlsafe_b64encode(url.strip().encode("utf-8")).decode("ascii").rstrip("=")
        share_id = "u!" + token
        data = self._get(f"/shares/{share_id}/driveItem")
        parent = data.get("parentReference", {}) or {}
        return {
            "item_id": data.get("id"),
            "drive_id": parent.get("driveId"),
            "site_id": parent.get("siteId"),
            "name": data.get("name"),
            "is_folder": "folder" in data,
            "web_url": data.get("webUrl"),
        }

    def resolve_site_by_url(self, site_url: str) -> str:
        """Resolve a SharePoint site URL (e.g. https://host/sites/Finance) to a site id."""
        p = urlparse(site_url.strip())
        host = p.netloc
        path = p.path.rstrip("/")
        endpoint = f"/sites/{host}:{path}" if path else f"/sites/{host}"
        data = self._get(endpoint)
        site_id = data.get("id")
        if not site_id:
            raise GraphError(f"could not resolve site from URL: {site_url}")
        return site_id

    # -- SharePoint lists (settings store) --------------------------------
    def find_list(self, site_id: str, display_name: str) -> str | None:
        data = self._get(f"/sites/{site_id}/lists",
                         params={"$select": "id,displayName,name"})
        for lst in data.get("value", []):
            if (lst.get("displayName") or lst.get("name")) == display_name:
                return lst.get("id")
        return None

    def create_list(self, site_id: str, display_name: str) -> str:
        """Create a simple key/value list (Title = key, Value = multiline text)."""
        body = {
            "displayName": display_name,
            "columns": [{"name": "Value", "text": {"allowMultipleLines": True}}],
            "list": {"template": "genericList"},
        }
        data = self._post(f"/sites/{site_id}/lists", body)
        return data["id"]

    def ensure_list(self, site_id: str, display_name: str) -> str:
        existing = self.find_list(site_id, display_name)
        return existing if existing else self.create_list(site_id, display_name)

    def list_columns(self, site_id: str, list_id: str) -> list[str]:
        """Return the internal column names of a list (used to filter fields we write)."""
        data = self._get(f"/sites/{site_id}/lists/{list_id}/columns",
                         params={"$select": "name,displayName"})
        return [c.get("name") for c in data.get("value", []) if c.get("name")]

    def list_items(self, site_id: str, list_id: str) -> list[dict]:
        data = self._get(f"/sites/{site_id}/lists/{list_id}/items",
                         params={"$expand": "fields", "$top": "500"})
        out = []
        for it in data.get("value", []):
            f = it.get("fields", {}) or {}
            out.append({"id": it.get("id"), "key": f.get("Title"), "value": f.get("Value")})
        return out

    def add_list_item(self, site_id: str, list_id: str, key: str, value: str) -> dict:
        return self._post(f"/sites/{site_id}/lists/{list_id}/items",
                          {"fields": {"Title": key, "Value": value}})

    def create_list_item_fields(self, site_id: str, list_id: str, fields: dict) -> dict:
        """Create a list item with an arbitrary set of column fields (used for the run log)."""
        return self._post(f"/sites/{site_id}/lists/{list_id}/items", {"fields": fields})

    def update_list_item(self, site_id: str, list_id: str, item_id: str, value: str) -> dict:
        return self._patch(f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
                           {"Value": value})

    def list_drives(self, site_id: str) -> list[dict]:
        data = self._get(f"/sites/{site_id}/drives")
        return [{"id": d.get("id"), "name": d.get("name"),
                 "driveType": d.get("driveType")} for d in data.get("value", [])]

    def list_children(self, drive_id: str, item_id: str | None = None) -> list[dict]:
        """List folders + files under a drive root (item_id=None) or a folder item."""
        if item_id:
            path = f"/drives/{drive_id}/items/{item_id}/children"
        else:
            path = f"/drives/{drive_id}/root/children"
        data = self._get(path, params={"$top": "200"})
        out = []
        for it in data.get("value", []):
            pr = it.get("parentReference", {}) or {}
            raw = pr.get("path", "") or ""
            folder = unquote(raw.split("root:", 1)[-1]) if "root:" in raw else unquote(raw)
            out.append({
                "id": it.get("id"),
                "name": it.get("name"),
                "is_folder": "folder" in it,
                "size": it.get("size"),
                "webUrl": it.get("webUrl"),
                "last_modified": it.get("lastModifiedDateTime"),
                "folder_path": folder,                       # e.g. /2026/EMEA/055/P05
            })
        return out

    def walk_files(self, drive_id: str, root_item_id: str | None = None, on_progress=None,
                   keep_folder=None, _depth: int = 0, max_depth: int = 12,
                   _state: dict | None = None) -> list[dict]:
        """Recursively list all files under a folder, skipping OLD/archive folders.

        keep_folder(name) -> bool lets the caller PRUNE whole branches (e.g. year/period folders
        that don't match the run), so we don't descend the entire tree.
        on_progress(state) is called as folders are scanned (state = {folders, files}).
        """
        from ..core.naming import is_ignored_dir
        if _state is None:
            _state = {"folders": 0, "files": 0}
        results: list[dict] = []
        try:
            children = self.list_children(drive_id, root_item_id)
        except GraphError:
            return results
        for it in children:
            if it["is_folder"]:
                if is_ignored_dir(it["name"]) or _depth >= max_depth:
                    continue
                if keep_folder is not None and not keep_folder(it["name"]):
                    continue                       # prune this branch (wrong year/period)
                _state["folders"] += 1
                if on_progress:
                    on_progress(_state)
                results.extend(self.walk_files(drive_id, it["id"], on_progress, keep_folder,
                                               _depth + 1, max_depth, _state))
            else:
                _state["files"] += 1
                results.append(it)
        return results

    def download_item(self, drive_id: str, item_id: str) -> bytes:
        url = f"{self._base}/drives/{drive_id}/items/{item_id}/content"
        with httpx.Client(timeout=120, follow_redirects=True) as c:
            r = c.get(url, headers=self._headers())
        if r.status_code >= 400:
            raise GraphError(f"download {item_id} -> {r.status_code}: {r.text[:300]}")
        return r.content

    def get_item_by_path(self, drive_id: str, item_path: str) -> dict | None:
        path = item_path.strip("/")
        url = f"{self._base}/drives/{drive_id}/root:/{path}"
        with httpx.Client(timeout=60) as c:
            r = c.get(url, headers=self._headers())
        if r.status_code == 404:
            return None
        if r.status_code >= 400:
            raise GraphError(f"get_item_by_path {path} -> {r.status_code}: {r.text[:300]}")
        return r.json()

    def download_by_path(self, drive_id: str, item_path: str) -> bytes | None:
        path = item_path.strip("/")
        url = f"{self._base}/drives/{drive_id}/root:/{path}:/content"
        with httpx.Client(timeout=120, follow_redirects=True) as c:
            r = c.get(url, headers=self._headers())
        if r.status_code == 404:
            return None
        if r.status_code >= 400:
            raise GraphError(f"download_by_path {path} -> {r.status_code}: {r.text[:300]}")
        return r.content

    # -- uploads -----------------------------------------------------------
    def upload_to_folder(self, drive_id: str, parent_item_id: str, filename: str,
                         data: bytes) -> dict:
        """Upload bytes as a new/updated file under a parent folder item."""
        if len(data) < _UPLOAD_SIMPLE_LIMIT:
            url = (f"{self._base}/drives/{drive_id}/items/{parent_item_id}:/"
                   f"{filename}:/content")
            with httpx.Client(timeout=120) as c:
                r = c.put(url, headers=self._headers(
                    {"Content-Type": "application/octet-stream"}), content=data)
            if r.status_code >= 400:
                raise GraphError(f"upload {filename} -> {r.status_code}: {r.text[:300]}")
            return r.json()
        return self._upload_large(drive_id, parent_item_id, filename, data)

    def upload_by_path(self, drive_id: str, item_path: str, data: bytes) -> dict:
        """Upload/replace a file addressed by path (used for settings.xlsx)."""
        path = item_path.strip("/")
        url = f"{self._base}/drives/{drive_id}/root:/{path}:/content"
        with httpx.Client(timeout=120) as c:
            r = c.put(url, headers=self._headers(
                {"Content-Type": "application/octet-stream"}), content=data)
        if r.status_code >= 400:
            raise GraphError(f"upload_by_path {path} -> {r.status_code}: {r.text[:300]}")
        return r.json()

    def _upload_large(self, drive_id: str, parent_item_id: str, filename: str,
                      data: bytes) -> dict:
        create = (f"{self._base}/drives/{drive_id}/items/{parent_item_id}:/"
                  f"{filename}:/createUploadSession")
        with httpx.Client(timeout=120) as c:
            r = c.post(create, headers=self._headers(),
                       json={"item": {"@microsoft.graph.conflictBehavior": "replace"}})
            if r.status_code >= 400:
                raise GraphError(f"createUploadSession -> {r.status_code}: {r.text[:300]}")
            upload_url = r.json()["uploadUrl"]

            chunk = 5 * 1024 * 1024
            total = len(data)
            start = 0
            last: dict = {}
            while start < total:
                end = min(start + chunk, total) - 1
                piece = data[start:end + 1]
                headers = {
                    "Content-Length": str(len(piece)),
                    "Content-Range": f"bytes {start}-{end}/{total}",
                }
                rr = c.put(upload_url, headers=headers, content=piece)
                if rr.status_code >= 400:
                    raise GraphError(f"upload chunk -> {rr.status_code}: {rr.text[:300]}")
                if rr.status_code in (200, 201):
                    last = rr.json()
                start = end + 1
            return last
