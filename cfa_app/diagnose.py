"""Outbound / Microsoft Graph connectivity check.

Run from the cfa_app folder:   python diagnose.py

Verifies this host can reach Microsoft login + Graph using the SAME stack the app uses (httpx,
which honours http(s)_proxy env vars just like the app). If .env is filled in, it also attempts a
real app-only token so you know auth works end to end. Handy on PythonAnywhere to confirm outbound
is allowed BEFORE configuring the web app.
"""

from __future__ import annotations

import os
import sys

import httpx


def _check(url: str, label: str) -> bool:
    try:
        r = httpx.get(url, timeout=20, follow_redirects=True)
        print(f"[OK]   {label}: HTTP {r.status_code}")
        return 200 <= r.status_code < 500
    except Exception as e:  # noqa: BLE001
        print(f"[FAIL] {label}: {type(e).__name__}: {e}")
        return False


def main() -> None:
    proxy = {k: v for k, v in os.environ.items() if "proxy" in k.lower()}
    print("Proxy env:", proxy or "(none — outbound is unrestricted)")
    print("-" * 64)

    reach = _check(
        "https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration",
        "login.microsoftonline.com",
    )
    reach = _check("https://graph.microsoft.com/v1.0/$metadata", "graph.microsoft.com") and reach
    print("-" * 64)
    print("Outbound to Graph:", "ALLOWED" if reach else "BLOCKED (paid plan needed)")
    print("-" * 64)

    # Optional end-to-end auth test if credentials are configured.
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    try:
        from app.config import get_settings
        settings = get_settings()
    except Exception as e:  # noqa: BLE001
        print("(could not load app settings:", e, ")")
        return

    if not settings.graph_configured:
        print("Graph credentials not set in .env — skipping token test.")
        print("(The reachability result above is what tells you if outbound is allowed.)")
        return

    try:
        from app.graph.auth import TokenProvider
        token = TokenProvider(settings).get_token()
        print(f"[OK]   acquired app-only token (length {len(token)}) - auth works end to end")
    except Exception as e:  # noqa: BLE001
        print(f"[FAIL] token acquisition: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
