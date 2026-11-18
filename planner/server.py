"""YD Ascend · Sales & Dispatch Planner - API + web UI.

Run:  python -m uvicorn server:app --app-dir planner --port 8600     (or double-click 'Start Planner.bat')
"""
from __future__ import annotations

import datetime as dt
import glob
import io
import os
import re
import threading
import uuid

import numpy as np
import openpyxl
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

import datasource as DS
import engine as E
from sales import SalesBook

HERE = os.path.dirname(os.path.abspath(__file__))
WEB = os.path.join(HERE, "web")
UPLOADS = os.path.join(HERE, "cache", "uploads")
os.makedirs(UPLOADS, exist_ok=True)

app = FastAPI(title="YD Ascend Sales & Dispatch Planner")
app.mount("/assets", StaticFiles(directory=os.path.join(WEB, "assets")), name="assets")

_lock = threading.Lock()
_cache: dict = {}


# ----------------------------------------------------------------------------- knowledge (auto re-sync)

def _knowledge_folder() -> str:
    return DS.resolve(DS.load_config().get("knowledge_folder", ".."))


def _fingerprint(sales: str) -> tuple:
    kf = _knowledge_folder()
    files = [sales] + sorted(glob.glob(os.path.join(kf, "*planning*.xlsx"))) + sorted(glob.glob(os.path.join(kf, "Priority Index*.xlsx")))
    return tuple((f, os.path.getmtime(f)) for f in files if os.path.exists(f))


def state(force: bool = False):
    with _lock:
        path, label = DS.sales_path(force_refresh=force)
        if not os.path.exists(path):
            raise HTTPException(500, f"Sales workbook not found: {path}")
        fp = _fingerprint(path)
        if force or _cache.get("fp") != fp:
            cfg = DS.load_config()
            pi = DS.resolve(cfg["priority_index"]) if cfg.get("priority_index") else None
            _cache.update(fp=fp, kb=E.load_knowledge(path, _knowledge_folder(), priority_index=pi),
                          book=SalesBook(path), label=label, path=path, loaded=dt.datetime.now())
        return _cache


# ----------------------------------------------------------------------------- pages

@app.get("/")
def page_dashboard():
    return FileResponse(os.path.join(WEB, "index.html"))


@app.get("/plan")
def page_plan():
    return FileResponse(os.path.join(WEB, "plan.html"))


# ----------------------------------------------------------------------------- status & sales views

@app.get("/api/status")
def api_status(refresh: bool = False):
    s = state(force=refresh)
    kb = s["kb"]
    return {
        "source": s["label"], "file": os.path.basename(s["path"]),
        "saved": dt.datetime.fromtimestamp(os.path.getmtime(s["path"])).strftime("%d %b %Y %H:%M"),
        "loaded": s["loaded"].strftime("%d %b %Y %H:%M"),
        **kb.summary(), "priority_index_rows": int(len(kb.pi)),
        "families": E.FAMILY_NAMES,
    }


@app.get("/api/dashboard")
def api_dashboard(fy: int | None = None):
    b = state()["book"]
    out = b.dashboard(fy)
    out["fys"] = b.fys()
    return out


@app.get("/api/projects")
def api_projects():
    return state()["book"].projects()


@app.get("/api/project")
def api_project(name: str):
    return state()["book"].project(name)


@app.get("/api/month")
def api_month(key: str):
    return state()["book"].month(key)


@app.get("/api/attention")
def api_attention():
    return state()["book"].attention()


@app.get("/api/cashflow")
def api_cashflow():
    return state()["book"].cashflow()


@app.get("/api/lookup")
def api_lookup(q: str):
    return state()["book"].lookup(q) if q.strip() else []


# ----------------------------------------------------------------------------- LoE upload

@app.post("/api/loe")
async def api_loe(file: UploadFile = File(...), sheet: str | None = None):
    data = await file.read()
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Upload the LoE as an .xlsx file.")
    token = uuid.uuid4().hex
    with open(os.path.join(UPLOADS, token + ".xlsx"), "wb") as fh:
        fh.write(data)
    with open(os.path.join(UPLOADS, token + ".name"), "w", encoding="utf-8") as fh:
        fh.write(file.filename)
    return _loe_info(token, sheet)


@app.get("/api/loe/{token}")
def api_loe_sheet(token: str, sheet: str | None = None):
    return _loe_info(token, sheet)


def _loe_path(token: str) -> str:
    if not re.fullmatch(r"[0-9a-f]{32}", token):
        raise HTTPException(400, "bad token")
    p = os.path.join(UPLOADS, token + ".xlsx")
    if not os.path.exists(p):
        raise HTTPException(404, "LoE not found - upload it again.")
    return p


def _loe_info(token: str, sheet: str | None):
    p = _loe_path(token)
    sheets = openpyxl.load_workbook(p, read_only=True).sheetnames
    sheet = sheet if sheet in sheets else sheets[0]
    try:
        loe = E.parse_loe(p, sheet)
    except Exception as e:  # noqa: BLE001
        raise HTTPException(422, f"Couldn't read this LoE: {e}") from e
    name = open(os.path.join(UPLOADS, token + ".name"), encoding="utf-8").read()
    txt = " ".join(str(v) for v in loe.info.values()) + " " + name
    pno = re.search(r"\b(\d{4})\b", str(loe.info.get("PROJECT NUMBER", ""))) or re.search(r"(\d{4})", name)
    client = re.sub(r"(?i)^m/s\.?\s*", "", str(loe.info.get("CLIENT", ""))).split(",")[0].strip()
    r = loe.rows
    return {
        "token": token, "filename": name, "sheet": sheet, "sheets": sheets, "info": loe.info,
        "rows": int(len(r)), "bought_out_rows": int(r.boi_list.sum()),
        "guess": {"project_no": pno[1] if pno else "", "client": client, "family": E.family_of(txt)},
        "preview": [{"sr": x.sr, "section": x.section, "name": x["name"], "drawing": x.drawing,
                     "qty": x.qty_txt or ("" if pd.isna(x.qty) else x.qty)} for _, x in r.head(400).iterrows()],
    }


# ----------------------------------------------------------------------------- planning

class PlanReq(BaseModel):
    token: str
    sheet: str | None = None
    project: str
    project_no: str = ""
    client: str = ""
    family: str
    pm: str = ""
    po_no: str = ""
    po_value: float
    currency: str = "INR"
    fx: float = 95.0
    billable_pct: float = 100.0
    po_date: dt.date
    delivery_date: dt.date
    start_month: dt.date | None = None
    end_month: dt.date | None = None
    mode: str = "assembly"          # assembly | lots
    lots: int = 3
    payment_terms: str = ""


def _po(req: PlanReq) -> E.POInfo:
    po = E.POInfo(project=req.project, family=req.family, po_value=req.po_value, currency=req.currency, fx=req.fx,
                  billable_pct=req.billable_pct, po_date=req.po_date, delivery_date=req.delivery_date, pm=req.pm)
    s, e = E.default_window(po)
    po.start_month = req.start_month or s.date()
    po.end_month = req.end_month or e.date()
    return po


def _clean(v):
    if v is None:
        return None
    if isinstance(v, (np.floating, float)):
        return None if np.isnan(v) else float(v)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.bool_,)):
        return bool(v)
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.strftime("%Y-%m")
    return v


@app.post("/api/plan")
def api_plan(req: PlanReq):
    if req.po_value <= 0:
        raise HTTPException(400, "Enter the PO value.")
    if req.delivery_date < req.po_date:
        raise HTTPException(400, "Delivery date is before the PO date.")
    kb = state()["kb"]
    po = _po(req)
    if pd.Timestamp(po.end_month) < pd.Timestamp(po.start_month):
        raise HTTPException(400, "Last dispatch month is before the first.")
    if req.mode == "lots":
        df = E.plan_lots(po, req.lots)
        df["drawing"] = ""
        df["category"] = "3/4"
        df["category_basis"] = "lot-wise plan"
        df["parent"] = ""
        value_mode = "equal lots"
    else:
        df = E.plan(E.parse_loe(_loe_path(req.token), req.sheet), po, kb)
        value_mode = df.attrs.get("value_mode", "")
    cols = ["row", "sr", "section", "name", "drawing", "qty", "billable", "billing_value", "dispatch_month",
            "dispatch_week", "category", "category_basis", "type", "reason", "basis", "match_score", "level",
            "has_children", "parent"]
    rows = []
    for k, (_, r) in enumerate(df.iterrows()):
        rows.append({c: _clean(r.get(c)) for c in cols} | {"idx": k})
    # month choices the user can pick from: PO month .. window end + 6
    first = E.month_start(min(pd.Timestamp(po.po_date), pd.Timestamp(po.start_month)))
    last = E.add_months(po.end_month, 6)
    months, m = [], first
    while m <= last:
        months.append({"key": m.strftime("%Y-%m"), "label": m.strftime("%b %Y"), "week": E.week_label(m)})
        m = E.add_months(m, 1)
    return {
        "rows": rows, "months": months, "target": po.target_inr, "value_mode": value_mode,
        "window": {"start": pd.Timestamp(po.start_month).strftime("%b %Y"), "end": pd.Timestamp(po.end_month).strftime("%b %Y")},
        "po": {**req.model_dump(mode="json"), "start_month": str(po.start_month), "end_month": str(po.end_month)},
    }


# ----------------------------------------------------------------------------- exports

class ExportReq(BaseModel):
    po: dict
    rows: list[dict]
    token: str | None = None
    sheet: str | None = None
    overwrite: bool = False


INDIAN = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'
NAVY = "2A5071"


def _rows_df(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df["billing_value"] = pd.to_numeric(df.billing_value, errors="coerce").fillna(0.0)
    df["billable"] = df.billable.astype(bool) & (df.billing_value > 0)
    df["dispatch_month"] = pd.to_datetime(df.dispatch_month.where(df.billable), format="%Y-%m", errors="coerce")
    df["dispatch_week"] = [E.week_label(m) if b and pd.notna(m) else "" for m, b in zip(df.dispatch_month, df.billable)]
    return df


def _style_table(ws, hdr_row, n, widths, money_cols):
    thin = Side(style="thin", color="D5DBE3")
    for c in ws[hdr_row]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=hdr_row + 1, max_row=hdr_row + n):
        for c in row:
            c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
            c.alignment = Alignment(vertical="center", wrap_text=True)
        for mc in money_cols:
            row[mc - 1].number_format = INDIAN
    for i, w in enumerate(widths):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w


@app.post("/api/export/plan")
def api_export_plan(req: ExportReq):
    po, df = req.po, _rows_df(req.rows)
    pno = po.get("project_no") or po.get("project", "plan")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Add Plan"
    ws.append(["Add Plan — Dispatch Week, Billing Value & Category"])
    ws.append([f"Project: {po.get('project')} · {len(df)} assemblies · PO value {po.get('currency')} "
               f"{po.get('po_value'):,.0f} · generated {dt.datetime.now():%d %b %Y %H:%M}"])
    ws.append([])
    ws.append(["SR No", "Assembly", "Drawing /Doc .No", "Dispatch Week", "Billing Value (INR)", "Category", "Basis"])
    for _, r in df.iterrows():
        ws.append([r.sr, r["name"], r.get("drawing") or "", r.dispatch_week,
                   float(r.billing_value) if r.billable else 0, (r.category if r.billable else "—") or "",
                   r.get("basis") or r.get("reason") or ""])
    ws.append([])
    ws.append(["", "TOTAL", "", "", f"=SUM(E5:E{4 + len(df)})"])
    ws["A1"].font = Font(bold=True, size=13, color=NAVY)
    ws["A2"].font = Font(color="666666")
    ws.cell(6 + len(df), 5).number_format = INDIAN
    ws.cell(6 + len(df), 2).font = Font(bold=True)
    _style_table(ws, 4, len(df), [7, 58, 24, 14, 18, 10, 46], [5])
    ws.freeze_panes = "A5"

    s = wb.create_sheet("Sales file rows")
    hdr = ["PROJECT", "BOI /MANF.", "PM", "ASSEMBLY", "RISK\nFB/ PAYMENT/ RED FLAG", "BILLING \nVALUE ($)",
           "BILLING \nVALUE", "DISP. \nSTATUS", "DISPATCH MONTH", "REMARKS"]
    s.append(hdr)
    v = df[df.billable]
    for _, r in v.iterrows():
        s.append([po.get("project"), "BOI" if r.get("type") == "BOI" else "", po.get("pm", ""), r["name"], "", "",
                  float(r.billing_value), "Not Disp.", E.sales_date(r.dispatch_month), ""])
    _style_table(s, 1, len(v), [18, 10, 16, 58, 12, 10, 16, 11, 15, 20], [7])
    for row in s.iter_rows(min_row=2, max_row=1 + len(v)):
        row[8].number_format = "dd-mmm-yy"

    m = wb.create_sheet("Month-wise")
    m.append(["Month", "Week", "Assemblies", "Billing Value (INR)"])
    g = v.groupby("dispatch_month").agg(n=("name", "size"), val=("billing_value", "sum")).reset_index()
    for _, r in g.iterrows():
        m.append([r.dispatch_month.strftime("%b %Y"), E.week_label(r.dispatch_month), int(r.n), float(r.val)])
    _style_table(m, 1, len(g), [12, 12, 12, 20], [4])

    p = wb.create_sheet("PO inputs")
    for k in ["project", "project_no", "client", "family", "pm", "po_no", "po_value", "currency", "fx", "billable_pct",
              "po_date", "delivery_date", "start_month", "end_month", "mode", "payment_terms"]:
        p.append([k.replace("_", " ").title(), str(po.get(k, ""))])
    p.column_dimensions["A"].width, p.column_dimensions["B"].width = 18, 50
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    fn = re.sub(r"[^\w\- ]+", "_", f"{pno} Add Plan.xlsx")
    return StreamingResponse(out, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f'attachment; filename="{fn}"'})


def _loe_filled(req: ExportReq) -> bytes:
    if not req.token:
        raise HTTPException(400, "No LoE for this plan (lot-wise plans have no LoE).")
    with open(_loe_path(req.token), "rb") as fh:
        data = fh.read()
    df = _rows_df(req.rows)
    df["billable"] = df.billable.astype(bool)
    return E.write_plan_into_loe(data, req.sheet, df)


@app.post("/api/export/loe")
def api_export_loe(req: ExportReq):
    pno = req.po.get("project_no") or "plan"
    return StreamingResponse(io.BytesIO(_loe_filled(req)),
                             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f'attachment; filename="{pno} planning.xlsx"'})


@app.post("/api/reference")
def api_reference(req: ExportReq):
    """Save the reviewed plan as '<pno> planning.xlsx' in the knowledge folder - future plans learn from it."""
    pno = re.sub(r"\D", "", str(req.po.get("project_no") or ""))[:6]
    if not pno:
        raise HTTPException(400, "Project number is needed to save a reference plan.")
    dest = os.path.join(_knowledge_folder(), f"{pno} planning.xlsx")
    if os.path.exists(dest) and not req.overwrite:
        return JSONResponse({"exists": True, "file": os.path.basename(dest)}, status_code=409)
    data = _loe_filled(req)
    with open(dest, "wb") as fh:
        fh.write(data)
    state(force=True)
    return {"saved": os.path.basename(dest)}
