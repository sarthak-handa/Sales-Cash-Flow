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
import urllib.error
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
    "sap_source": {"type": "local", "path": "../revenue data sap.xlsx"},
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


SECRETS_PATH = os.path.join(HERE, "secrets.json")   # git-ignored: the repo is public
BILLING_DOC_TYPES = ("F2", "S1", "G2", "S2")           # invoice, cancellation, credit memo, credit-memo cancellation
_status: dict = {"sales": "", "sap": "", "sales_failed_at": 0.0}


def load_secrets() -> dict:
    if not os.path.exists(SECRETS_PATH):
        return {}
    with open(SECRETS_PATH, encoding="utf-8") as fh:
        return json.load(fh)


def _graph_cfg(cfg: dict) -> dict:
    g = dict(cfg.get("graph", {}))
    for k, v in load_secrets().get("graph", {}).items():   # secrets.json fills what config.json leaves blank
        if v and not g.get(k):
            g[k] = v
    if not g.get("client_secret"):
        g["client_secret"] = os.environ.get(g.get("client_secret_env", ""), "")
    return g


def _graph_token(g: dict) -> str:
    if not (g.get("tenant_id") and g.get("client_id") and g.get("client_secret")):
        raise RuntimeError("Graph: tenant_id / client_id / client_secret missing (planner/secrets.json).")
    data = urllib.parse.urlencode({
        "client_id": g["client_id"], "client_secret": g["client_secret"], "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }).encode()
    url = f"https://login.microsoftonline.com/{g['tenant_id']}/oauth2/v2.0/token"
    try:
        with urllib.request.urlopen(urllib.request.Request(url, data=data), timeout=30) as r:
            return json.load(r)["access_token"]
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        if "AADSTS7000222" in body:
            raise RuntimeError("the Azure app's client secret has expired; IT must create a new one "
                               "(put it in planner/secrets.json)") from None
        raise RuntimeError(f"Microsoft sign-in failed ({e.code}): {body[:200]}") from None


def _graph_download(g: dict, dest: str) -> None:
    token = _graph_token(g)
    url = f"https://graph.microsoft.com/v1.0/drives/{g['drive_id']}/items/{g['item_id']}/content"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    tmp = dest + ".part"
    with urllib.request.urlopen(req, timeout=120) as r, open(tmp, "wb") as fh:
        fh.write(r.read())
    os.replace(tmp, dest)


def _stale(path: str, minutes: float) -> bool:
    return not os.path.exists(path) or time.time() - os.path.getmtime(path) > 60 * float(minutes)


def sales_path(force_refresh: bool = False) -> tuple[str, str]:
    """(path to an .xlsx to read, human description of the source).
    graph: the live workbook, cached; if the download fails, the last download, else the local file."""
    cfg = load_config()
    src = cfg["sales_source"]
    local = resolve(src.get("path", DEFAULT_CONFIG["sales_source"]["path"]))
    if src.get("type") != "graph":
        return local, "Local file"
    g = _graph_cfg(cfg)
    minutes = float(g.get("refresh_minutes", 15))
    dest = resolve(g.get("cache_path", "cache/sales_live.xlsx"))
    recently_failed = time.time() - _status["sales_failed_at"] < 60 * minutes
    if (force_refresh or _stale(dest, minutes)) and (force_refresh or not recently_failed):
        try:
            _graph_download(g, dest)
            _status["sales"], _status["sales_failed_at"] = "", 0.0
        except Exception as e:                           # keep working on the last good copy
            _status["sales"], _status["sales_failed_at"] = str(e), time.time()
    if _status["sales"]:
        if os.path.exists(dest):
            return dest, f"Live workbook, last download (sync failing: {_status['sales']})"
        return local, f"Local file (live sync failing: {_status['sales']})"
    return dest, "Live workbook (SharePoint)"


# ----------------------------------------------------------------------------- SAP billing

def _sap_fetch(s: dict, dest: str) -> None:
    """Pull every row of the SAP billing OData service, keep the billing documents, save as a pickle."""
    import base64

    import pandas as pd
    auth = base64.b64encode(f"{s['username']}:{s['password']}".encode()).decode()
    rows, skip = [], 0
    while True:
        url = s["url"] + f"?$format=json&$top=5000&$skip={skip}"
        req = urllib.request.Request(url, headers={"Authorization": "Basic " + auth, "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                d = json.load(r)["d"]
        except urllib.error.HTTPError as e:
            raise RuntimeError(f"SAP API error {e.code}" + (" - check the user name / password in planner/secrets.json"
                                                             if e.code in (401, 403) else "")) from None
        res = d["results"] if isinstance(d, dict) else d
        rows += res
        if len(res) < 5000:
            break
        skip += 5000
    df = pd.DataFrame(rows).drop(columns=["__metadata", "Parameters"], errors="ignore")
    df = df[df["doc_type"].astype(str).str.strip().isin(BILLING_DOC_TYPES)].copy()
    for c in ("Billing_date", "so_dt", "p_from_dt", "p_to_dt"):            # OData "/Date(ms)/"
        if c in df:
            ms = pd.to_numeric(df[c].astype(str).str.extract(r"(-?\d+)")[0], errors="coerce")
            df[c] = pd.to_datetime(ms, unit="ms")
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    df.to_pickle(dest + ".part")
    os.replace(dest + ".part", dest)


def sap_path(force_refresh: bool = False) -> str:
    """Path to the SAP billing data: the live API's cached pull (.pkl), or the Excel export (sheet Query1)."""
    cfg = load_config()
    src = cfg.get("sap_source", DEFAULT_CONFIG["sap_source"])
    local = resolve(src.get("path", "../revenue data sap.xlsx"))
    if src.get("type") != "api":
        return local
    s = load_secrets().get("sap", {})
    dest = resolve(src.get("cache_path", "cache/sap_live.pkl"))
    if s.get("url") and (force_refresh or _stale(dest, src.get("refresh_minutes", 30))):
        try:
            _sap_fetch(s, dest)
            _status["sap"] = ""
        except Exception as e:
            _status["sap"] = str(e)
            if os.path.exists(dest):
                os.utime(dest)                          # retry after refresh_minutes, not on every request
    return dest if os.path.exists(dest) else local


def sap_label() -> str:
    if load_config().get("sap_source", {}).get("type") == "api":
        return "SAP live (S/4HANA)" + (f" - last pull failed: {_status['sap']}" if _status["sap"] else "")
    return "SAP export file"


def status() -> dict:
    return dict(_status)
