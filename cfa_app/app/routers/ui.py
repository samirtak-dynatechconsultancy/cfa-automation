"""Server-rendered UI pages (Jinja2): run page and admin page."""

from __future__ import annotations

import os
from pathlib import Path

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from ..deps import app_settings, get_settings_store

router = APIRouter(tags=["ui"])
_templates = Jinja2Templates(directory=str(Path(__file__).resolve().parent.parent / "templates"))
_STATIC_DIR = Path(__file__).resolve().parent.parent / "static"


def _asset_version() -> str:
    """A stamp that changes whenever app.js / style.css change, so a deploy busts the browser cache
    (important when embedded in a SharePoint iframe, which caches static assets aggressively)."""
    try:
        return str(int(max(os.path.getmtime(_STATIC_DIR / f) for f in ("app.js", "style.css"))))
    except OSError:
        return "1"


_templates.env.globals["asset_version"] = _asset_version


_NO_CACHE = {"Cache-Control": "no-cache, must-revalidate"}   # always revalidate the page HTML


@router.get("/", response_class=HTMLResponse)
def run_page(request: Request):
    store = get_settings_store()
    return _templates.TemplateResponse(
        request,
        "run.html",
        {
            "graph_configured": app_settings().graph_configured,
            "defaults": store.load(),
        },
        headers=_NO_CACHE,
    )


@router.get("/admin", response_class=HTMLResponse)
def admin_page(request: Request):
    store = get_settings_store()
    return _templates.TemplateResponse(
        request,
        "admin.html",
        {
            "graph_configured": app_settings().graph_configured,
            "persisted": store.persisted,
            "values": store.load(),
        },
        headers=_NO_CACHE,
    )
