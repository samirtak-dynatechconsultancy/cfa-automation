"""Report the deployed build/commit so it can be checked at runtime (e.g. GET /version).

Deploys can lag the branch, so this exposes exactly what code is running. It tries, in order:
the git checkout, a build-info file written at deploy time, then known CI/host env vars.
"""

from __future__ import annotations

import json
import os
import subprocess
from functools import lru_cache
from pathlib import Path

_REPO = Path(__file__).resolve().parents[2]     # repo root (…/cfa_app/app/version.py -> repo)


@lru_cache
def get_version_info() -> dict:
    info = {"commit": None, "commit_short": None, "subject": None, "commit_date": None,
            "branch": None, "source": "unknown"}

    # 1. Live git checkout (present when the deploy artifact includes .git).
    try:
        out = subprocess.run(
            ["git", "-C", str(_REPO), "log", "-1", "--format=%H%x1f%s%x1f%ci%x1f%D"],
            capture_output=True, text=True, timeout=5)
        if out.returncode == 0 and out.stdout.strip():
            parts = (out.stdout.strip().split("\x1f") + ["", "", "", ""])[:4]
            h, subject, date, refs = parts
            info.update(commit=h, commit_short=h[:7], subject=subject,
                        commit_date=date, branch=refs.strip(), source="git")
            return info
    except Exception:
        pass

    # 2. Build-info file written at deploy time (see the GitHub Actions workflow).
    for p in (_REPO / "_build_info.json", Path(__file__).with_name("_build_info.json")):
        try:
            if p.exists():
                d = json.loads(p.read_text(encoding="utf-8"))
                c = d.get("commit") or ""
                info.update(commit=c or None, commit_short=c[:7] or None,
                            subject=d.get("subject"), commit_date=d.get("date"),
                            branch=d.get("branch"), source="build-file")
                return info
        except Exception:
            pass

    # 3. Known CI / host environment variables.
    for var in ("APP_COMMIT", "SCM_COMMIT_ID", "GITHUB_SHA", "BUILD_COMMIT", "BUILD_SOURCEVERSION"):
        v = os.environ.get(var)
        if v:
            info.update(commit=v, commit_short=v[:7], source=f"env:{var}")
            return info

    return info
