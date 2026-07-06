"""Local launcher: `python run.py` (or use `uvicorn app.main:app --reload`)."""

import uvicorn

from app.config import get_settings
from app.logging_config import setup_logging

if __name__ == "__main__":
    s = get_settings()
    setup_logging(s.log_level)
    uvicorn.run("app.main:app", host=s.app_host, port=s.app_port,
                reload=True, log_level=s.log_level.lower())
