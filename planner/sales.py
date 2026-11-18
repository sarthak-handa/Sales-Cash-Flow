"""Sales & cash-flow views over the SALES & CASH FLOW workbook (the live status of every project)."""
from __future__ import annotations

import datetime as dt
import re

import numpy as np
import openpyxl
import pandas as pd

import engine as E


def fy_of(d) -> int:
    """Indian financial year label = year in which it starts (Apr-Mar). Nov 2026 -> 2026 (FY 2026-27)."""
    d = pd.Timestamp(d)
    return d.year if d.month >= 4 else d.year - 1


def fy_months(fy: int) -> list[pd.Timestamp]:
    return [pd.Timestamp(fy, 4, 1) + pd.DateOffset(months=k) for k in range(12)]


def _f(v):
    return None if v is None or (isinstance(v, float) and np.isnan(v)) else float(v)


class SalesBook:
    def __init__(self, path: str, today: dt.date | None = None):
        self.path = path
        self.df = E.load_sales(path)
        self.df["bill"] = self.df.bill.fillna(0.0)
        self.today = pd.Timestamp(today or dt.date.today())
        self.cur = E.month_start(self.today)
        d = self.df
        d["slipped"] = ~d.dispatched & (d.month < self.cur)
        d["upcoming"] = ~d.dispatched & (d.month >= self.cur)

    # ------------------------------------------------------------------ overview
    def fys(self) -> list[int]:
        m = self.df.month.dropna()
        return sorted({fy_of(x) for x in m})

    def dashboard(self, fy: int | None = None) -> dict:
        fy = fy if fy is not None else fy_of(self.cur)
        d = self.df
        months = fy_months(fy)
        in_fy = d.month.isin(months)
        f = d[in_fy]
        nxt = [E.add_months(self.cur, k) for k in range(3)]
        series = []
        for m in months:
            g = d[d.month == m]
            series.append({
                "month": m.strftime("%b %Y"), "key": m.strftime("%Y-%m"), "week": E.week_label(m),
                "dispatched": float(g.loc[g.dispatched, "bill"].sum()),
                "slipped": float(g.loc[g.slipped, "bill"].sum()),
                "planned": float(g.loc[g.upcoming, "bill"].sum()),
                "assemblies": int(len(g)),
                "is_current": bool(m == self.cur), "is_past": bool(m < self.cur),
            })
        quarters = []
        for q in range(4):
            qm = months[q * 3:(q + 1) * 3]
            g = d[d.month.isin(qm)]
            quarters.append({"label": f"Q{q + 1} ({qm[0]:%b}–{qm[-1]:%b})",
                             "total": float(g.bill.sum()),
                             "dispatched": float(g.loc[g.dispatched, "bill"].sum())})
        cm = d[d.month == self.cur]
        slip_all = d[d.slipped]
        return {
            "fy": fy, "fy_label": f"FY {fy}-{str(fy + 1)[-2:]}", "today": self.today.strftime("%d %b %Y"),
            "current_month": self.cur.strftime("%b %Y"),
            "kpi": {
                "fy_total": float(f.bill.sum()),
                "fy_dispatched": float(f.loc[f.dispatched, "bill"].sum()),
                "month_plan": float(cm.bill.sum()),
                "month_dispatched": float(cm.loc[cm.dispatched, "bill"].sum()),
                "slipped_value": float(slip_all.bill.sum()),
                "slipped_count": int(len(slip_all)),
                "next3": float(d[d.month.isin(nxt) & ~d.dispatched].bill.sum()),
                "next3_label": f"{nxt[0]:%b}–{nxt[-1]:%b %Y}",
                "projects_active": int(d[~d.dispatched & (d.bill > 0)].project.nunique()),
            },
            "series": series,
            "quarters": quarters,
        }

    # ------------------------------------------------------------------ tables
    def projects(self) -> list[dict]:
        d = self.df
        out = []
        for p, g in d.groupby("project"):
            pend = g[~g.dispatched & (g.bill > 0)]
            nextm = pend[pend.month >= self.cur].month.min()
            risks = g[~g.dispatched & (g.risk != "")]
            out.append({
                "project": p, "line": E.family_of(p), "pm": ", ".join(sorted({x for x in g.pm if x}))[:60],
                "assemblies": int(len(g)),
                "total": float(g.bill.sum()),
                "dispatched": float(g.loc[g.dispatched, "bill"].sum()),
                "pending": float(pend.bill.sum()),
                "slipped": float(g.loc[g.slipped, "bill"].sum()),
                "slipped_count": int(g.slipped.sum()),
                "next_month": nextm.strftime("%b %Y") if pd.notna(nextm) else "",
                "next_key": nextm.strftime("%Y-%m") if pd.notna(nextm) else "",
                "last_month": g.month.max().strftime("%b %Y") if g.month.notna().any() else "",
                "risks": int(len(risks)),
                "pct": float(g.loc[g.dispatched, "bill"].sum() / g.bill.sum()) if g.bill.sum() else 0.0,
            })
        return out

    def _rows(self, g: pd.DataFrame) -> list[dict]:
        return [{
            "project": r.project, "assembly": r["name"], "type": r.boi, "pm": r.pm, "risk": r.risk,
            "bill": float(r.bill), "bill_usd": _f(r.bill_usd),
            "month": r.month.strftime("%b %Y") if pd.notna(r.month) else "",
            "key": r.month.strftime("%Y-%m") if pd.notna(r.month) else "",
            "week": E.week_label(r.month) if pd.notna(r.month) else "",
            "status": "Dispatched" if r.dispatched else ("Slipped" if r.slipped else "Planned"),
            "remarks": r.remarks, "vendor": r.vendor,
        } for _, r in g.iterrows()]

    def project(self, name: str) -> list[dict]:
        return self._rows(self.df[self.df.project == name].sort_values(["month", "name"]))

    def month(self, key: str) -> list[dict]:
        m = pd.Timestamp(key + "-01")
        return self._rows(self.df[self.df.month == m].sort_values(["project", "name"]))

    def attention(self) -> list[dict]:
        """Slipped (past month, not dispatched) or flagged with a risk owner."""
        d = self.df
        g = d[(d.slipped | ((d.risk != "") & ~d.dispatched)) & (d.bill > 0)]
        return self._rows(g.sort_values(["month", "project"]))

    def cashflow(self) -> dict:
        """CASH FLOW sheet (pivot of payment requirement by month) as stored in the workbook."""
        try:
            wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
            ws = wb["CASH FLOW"]
        except Exception:  # noqa: BLE001
            return {"columns": [], "rows": []}
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        hi = next((i for i, r in enumerate(raw) if r and str(r[0] or "").strip().upper() == "ROW LABELS"), None)
        if hi is None:
            return {"columns": [], "rows": []}
        months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        hdr = list(raw[hi][1:])
        while hdr and hdr[-1] is None:
            hdr.pop()
        cols, prev = [], None
        for c in hdr:
            s = re.sub(r"\s+", " ", str(c or "")).strip()
            s = re.sub(r"(?i)^sum of payment req\.?", "", s).strip().upper()[:3]
            if s not in months and prev is not None:     # unlabeled column = next month in sequence
                s = months[(months.index(prev) + 1) % 12]
            prev = s if s in months else prev
            cols.append(s.title() if s else "—")
        rows = []
        for r in raw[hi + 1:]:
            if not r or r[0] is None or str(r[0]).strip().lower() == "(blank)":
                continue
            vals = [_f(v) if isinstance(v, (int, float)) else None for v in r[1:len(cols) + 1]]
            if not any(vals):
                continue
            rows.append({"project": str(r[0]).strip(), "values": vals, "total": float(sum(v or 0 for v in vals))})
        return {"columns": cols, "rows": rows}

    # ------------------------------------------------------------------ lookup
    def lookup(self, q: str, limit: int = 60) -> list[dict]:
        from rapidfuzz import fuzz, process
        d = self.df[self.df.bill > 0]
        keys = d.key.drop_duplicates().tolist()
        hits = process.extract(E.norm(q), keys, scorer=fuzz.token_set_ratio, limit=40, score_cutoff=70)
        g = d[d.key.isin({k for k, _, _ in hits})].sort_values("month", ascending=False).head(limit)
        return self._rows(g)
