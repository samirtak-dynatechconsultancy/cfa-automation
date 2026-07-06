"""SharePoint browsing API: search sites, list document libraries, list folders/files."""

from __future__ import annotations

from fastapi import APIRouter, HTTPException, Query
from pydantic import BaseModel

from ..deps import get_graph_client
from ..graph.client import GraphError

router = APIRouter(prefix="/api", tags=["browse"])


class ResolveBody(BaseModel):
    url: str
    expect: str = "any"      # "folder" | "file" | "any"


@router.post("/resolve")
def resolve_link(body: ResolveBody):
    if not body.url.strip():
        raise HTTPException(status_code=400, detail="Paste a SharePoint link.")
    try:
        item = get_graph_client().resolve_share_link(body.url)
    except GraphError:
        raise HTTPException(
            status_code=400,
            detail="Could not resolve that link. Check the link is a SharePoint 'Copy link' URL "
                   "and that the app has access to the site.",
        )
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))

    if not item.get("drive_id") or not item.get("item_id"):
        raise HTTPException(status_code=400, detail="Link did not resolve to a SharePoint item.")
    if body.expect == "folder" and not item["is_folder"]:
        raise HTTPException(status_code=400,
                            detail=f"That link points to a file ({item['name']}), not a folder.")
    if body.expect == "file" and item["is_folder"]:
        raise HTTPException(status_code=400,
                            detail=f"That link points to a folder ({item['name']}), not a file.")
    return item


@router.get("/sites")
def search_sites(q: str = Query("", description="site search text")):
    try:
        return {"sites": get_graph_client().search_sites(q)}
    except GraphError as e:
        raise HTTPException(status_code=502, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/drives")
def list_drives(site_id: str = Query(...)):
    try:
        return {"drives": get_graph_client().list_drives(site_id)}
    except GraphError as e:
        raise HTTPException(status_code=502, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.get("/items")
def list_items(drive_id: str = Query(...), item_id: str | None = Query(None)):
    try:
        return {"items": get_graph_client().list_children(drive_id, item_id)}
    except GraphError as e:
        raise HTTPException(status_code=502, detail=str(e))
    except RuntimeError as e:
        raise HTTPException(status_code=400, detail=str(e))
