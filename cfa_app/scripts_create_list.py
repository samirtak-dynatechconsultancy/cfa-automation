"""One-off: ensure the 'CFA App Settings' SharePoint List exists on the configured site.
Run from the cfa_app directory so .env is picked up:  python scripts_create_list.py
"""
import sys
sys.path.insert(0, ".")

from app.config import get_settings
from app.deps import get_graph_client

s = get_settings()
print(f"Site URL   : {s.settings_site_url or '(not set)'}")
print(f"List name  : {s.settings_list_name}")
print(f"Graph creds: {'configured' if s.graph_configured else 'MISSING'}")
if not s.graph_configured or not s.settings_site_url:
    raise SystemExit("Missing GRAPH_* or SETTINGS_SITE_URL in .env")

g = get_graph_client()

print("\n[1] Acquiring token + resolving site…")
site_id = g.resolve_site_by_url(s.settings_site_url)
print(f"    site id = {site_id}")

print("\n[2] Looking for existing list…")
existing = g.find_list(site_id, s.settings_list_name)
if existing:
    print(f"    already exists: {existing}")
    list_id = existing
else:
    print("    not found; creating…")
    list_id = g.create_list(site_id, s.settings_list_name)
    print(f"    CREATED: {list_id}")

items = g.list_items(site_id, list_id)
print(f"\n[3] List ready. Current item count: {len(items)}")
print("\nDONE — the app can now persist settings to this list.")
