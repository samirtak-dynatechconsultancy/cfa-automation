"""WSGI entrypoint for PythonAnywhere.

PythonAnywhere web apps are WSGI; this FastAPI app is ASGI, so we wrap it with a2wsgi.
The PythonAnywhere WSGI config file should add this folder to sys.path and then do:

    from wsgi import application

See DEPLOY_PYTHONANYWHERE.md for the full steps.
"""

from a2wsgi import ASGIMiddleware

from app.main import app as _asgi_app

application = ASGIMiddleware(_asgi_app)
