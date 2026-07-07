"""Run API: start a transfer run (async) and poll its status/result."""

from __future__ import annotations

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from ..deps import get_run_manager
from ..services.runs import RunParams

router = APIRouter(prefix="/api/runs", tags=["runs"])


class StartRun(BaseModel):
    year: int
    period_from: int
    period_to: int
    source_drive_id: str
    source_folder_id: str
    master_drive_id: str
    master_item_id: str
    master_name: str
    output_drive_id: str
    output_folder_id: str
    source_folder_name: str = ""
    output_folder_name: str = ""


@router.post("")
def start_run(body: StartRun):
    if not (1 <= body.period_from <= 12 and 1 <= body.period_to <= 12):
        raise HTTPException(status_code=400, detail="Periods must be between 1 and 12.")
    if body.period_from > body.period_to:
        raise HTTPException(status_code=400, detail="'Period from' must not be after 'Period to'.")
    params = RunParams(**body.model_dump())
    run_id = get_run_manager().start(params)
    return {"run_id": run_id}


@router.get("/latest")
def latest_run():
    """The most recent run (for re-attaching after a refresh). {} if there is none."""
    result = get_run_manager().latest()
    return result.as_dict() if result is not None else {}


@router.get("/{run_id}")
def get_run(run_id: str):
    result = get_run_manager().get(run_id)
    if result is None:
        raise HTTPException(status_code=404, detail="run not found")
    return result.as_dict()
