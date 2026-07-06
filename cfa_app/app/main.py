"""FastAPI application entrypoint for the CFA consistency-check transfer tool."""

from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from .config import get_settings
from .logging_config import get_logger, setup_logging
from .routers import runs, settings, sites, ui

setup_logging(get_settings().log_level)
get_logger("app").info("CFA transfer app starting (graph_configured=%s)",
                       get_settings().graph_configured)

app = FastAPI(title="CFA Consistency-Check Transfer", version="1.0.0")

_static_dir = Path(__file__).resolve().parent / "static"
app.mount("/static", StaticFiles(directory=str(_static_dir)), name="static")

@app.middleware("http")
async def _frame_headers(request, call_next):
    """Allow embedding in an <iframe> from the configured hosts (e.g. SharePoint)."""
    response = await call_next(request)
    ancestors = get_settings().frame_ancestors
    if ancestors:
        response.headers["Content-Security-Policy"] = f"frame-ancestors {ancestors}"
        if "x-frame-options" in response.headers:
            del response.headers["X-Frame-Options"]   # can't express multi-origin; CSP supersedes
    return response


app.include_router(ui.router)
app.include_router(sites.router)
app.include_router(runs.router)
app.include_router(settings.router)


@app.get("/health", tags=["ops"])
def health():
    return {"status": "ok"}
