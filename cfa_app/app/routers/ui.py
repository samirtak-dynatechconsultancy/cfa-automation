"""Server-rendered UI pages (Jinja2): run page and admin page."""

from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from ..deps import app_settings, get_settings_store

router = APIRouter(tags=["ui"])
_templates = Jinja2Templates(directory=str(Path(__file__).resolve().parent.parent / "templates"))


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
    )
