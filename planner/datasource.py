"""Where the SALES & CASH FLOW workbook comes from.

type "local": a file on disk (default - the OneDrive-synced copy next to this folder)
type "graph": the live workbook in OneDrive/SharePoint via Microsoft Graph. The workbook is
              downloaded to a cache file every `refresh_minutes`, so the rest of the system
              always reads a normal .xlsx. Needs an Azure AD app (client credentials) with
              Files.Read.All / Sites.Read.All application permission.

Edit planner/config.json to switch.
"""
from __future__ import annotations

import json
import os
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CONFIG_PATH = os.path.join(HERE, "config.json")

DEFAULT_CONFIG = {
    "sales_source": {
        "type": "local",
        "path": "../SALES & CASH FLOW.xlsx",
    },
    "graph": {
        "tenant_id": "",
        "client_id": "",
        "client_secret_env": "YD_GRAPH_CLIENT_SECRET",
        "drive_id": "",
        "item_id": "",
        "refresh_minutes": 15,
        "cache_path": "cache/sales_live.xlsx",
    },
    "knowledge_folder": "..",
    "priority_index": "",
    "reference_plans_folder": "..",
}


def load_config() -> dict:
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
            json.dump(DEFAULT_CONFIG, fh, indent=2)
    with open(CONFIG_PATH, encoding="utf-8") as fh:
        cfg = json.load(fh)
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    return cfg


def resolve(p: str) -> str:
    return p if os.path.isabs(p) else os.path.normpath(os.path.join(HERE, p))


def _graph_token(g: dict) -> str:
    secret = os.environ.get(g.get("client_secret_env", ""), "")
    if not (g.get("tenant_id") and g.get("client_id") and secret):
        raise RuntimeError("Graph source selected but tenant_id / client_id / client secret env var are not set.")
    data = urllib.parse.urlencode({
        "client_id": g["client_id"], "client_secret": secret, "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }).encode()
    url = f"https://login.microsoftonline.com/{g['tenant_id']}/oauth2/v2.0/token"
    with urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=30) as r:
        return json.load(r)["access_token"]


def _graph_download(g: dict, dest: str) -> None:
    token = _graph_token(g)
    url = f"https://graph.microsoft.com/v1.0/drives/{g['drive_id']}/items/{g['item_id']}/content"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp = dest + ".part"
    with urllib.request.urlopen(req, timeout=120) as r, open(tmp, "wb") as fh:
        fh.write(r.read())
    os.replace(tmp, dest)


def sales_path(force_refresh: bool = False) -> tuple[str, str]:
    """(path to an .xlsx to read, human description of the source)"""
    cfg = load_config()
    src = cfg["sales_source"]
    if src.get("type") == "graph":
        g = cfg["graph"]
        dest = resolve(g.get("cache_path", "cache/sales_live.xlsx"))
        stale = (not os.path.exists(dest)
                 or time.time() - os.path.getmtime(dest) > 60 * float(g.get("refresh_minutes", 15)))
        if force_refresh or stale:
            _graph_download(g, dest)
        return dest, "Live workbook (Microsoft Graph)"
    return resolve(src.get("path", DEFAULT_CONFIG["sales_source"]["path"])), "Local file"
