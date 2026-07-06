"""Central logging setup. Logs go to stdout under the 'cfa' logger namespace.

Level comes from LOG_LEVEL (INFO by default). Set LOG_LEVEL=DEBUG to also see every Graph
HTTP call. The 'cfa' logger uses its own handler and does not propagate, so it prints cleanly
alongside uvicorn's request log without duplication.
"""

from __future__ import annotations

import logging
import sys

_configured = False


def setup_logging(level: str = "INFO") -> None:
    global _configured
    if _configured:
        return
    logger = logging.getLogger("cfa")
    logger.setLevel(level.upper())
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter(
        fmt="%(asctime)s %(levelname)-7s %(name)s | %(message)s",
        datefmt="%H:%M:%S",
    ))
    logger.addHandler(handler)
    logger.propagate = False
    _configured = True


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger("cfa." + name)
