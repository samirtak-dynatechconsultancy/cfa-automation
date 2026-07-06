"""Settings API: read and update admin defaults + detection config (settings.xlsx)."""

from __future__ import annotations

from fastapi import APIRouter, HTTPException

from ..deps import get_settings_store
from ..graph.client import GraphError

router = APIRouter(prefix="/api/settings", tags=["settings"])


@router.get("")
def read_settings():
    store = get_settings_store()
    return {"persisted": store.persisted, "values": store.load()}


@router.put("")
def write_settings(values: dict):
    store = get_settings_store()
    try:
        saved = store.save(values)
    except GraphError as e:
        raise HTTPException(status_code=502, detail=f"could not save to SharePoint: {e}")
    return {"persisted": store.persisted, "values": saved}
