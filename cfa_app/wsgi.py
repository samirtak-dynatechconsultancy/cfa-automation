"""WSGI entrypoint for PythonAnywhere.

PythonAnywhere web apps are WSGI; this FastAPI app is ASGI, so we bridge it with a2wsgi.

IMPORTANT: a2wsgi's ASGIMiddleware starts a background event-loop thread. uWSGI imports the app in
the MASTER process and then FORKS worker processes to serve requests — and threads do not survive a
fork. If we built the bridge at import time (in the master), every request in a worker would submit
work to a loop whose thread no longer exists and hang until harakiri. So we build it LAZILY on the
first request, which runs inside the worker (post-fork).

The PythonAnywhere WSGI config file should add this folder to sys.path and then:

    from wsgi import application

See DEPLOY_PYTHONANYWHERE.md for the full steps.
"""

from __future__ import annotations

import threading

_app = None
_lock = threading.Lock()


def application(environ, start_response):
    global _app
    if _app is None:
        with _lock:
            if _app is None:
                from a2wsgi import ASGIMiddleware
                from app.main import app as asgi_app
                _app = ASGIMiddleware(asgi_app)
    return _app(environ, start_response)
