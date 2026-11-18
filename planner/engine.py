"""Dispatch & billing planner engine.

Learns from:
  * SALES & CASH FLOW.xlsx (PROJECTS sheet) - every past assembly with billing value and dispatch month
  * reference plans (*planning.xlsx)        - the planner's own LoEs, showing which rows he bills and which he doesn't

Given a new LoE + PO details it proposes, per row: billable or not, billing value, dispatch month/week.
"""
from __future__ import annotations

import calendar
import datetime as dt
import glob
import os
import re
from dataclasses import dataclass, field

import numpy as np
import openpyxl
import pandas as pd
from rapidfuzz import fuzz, process

# ----------------------------------------------------------------------------- text helpers

_STOP = {"ASSY", "ASSEMBLY", "ASSEMBLIES", "NOS", "NO", "SET", "SETS", "WITH", "AND", "OF", "THE", "FOR",
         "IN", "TO", "ON", "AT", "A", "DRG", "QTY", "TYPE"}


def norm(s) -> str:
    """Normalise an assembly name for matching: upper case, drop brackets, numbers, 'ASSY' etc."""
    s = str(s or "").upper()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"#\s*\d+", " ", s)
    s = re.sub(r"@.*", " ", s)  # '@ UNCOILER' location suffix
    s = re.sub(r"[^A-Z ]", " ", s)
    words = [w for w in s.split() if w not in _STOP and len(w) > 1]
    return " ".join(words)


FAMILIES = [  # (family, keywords) - first match wins
    ("ARP", ["ARP", "ACID REGEN"]),
    ("CCL", ["CCL", "COLOR COAT", "COLOUR COAT"]),
    ("CGL", ["CGL", "GI/GL", "GL LINE", "GALV", "AL-ZN", "ZN LINE", "GALVALUME", "NON-OX"]),
    ("PICKLING", ["PICKLING", "CPL", "PPL", "PUSH PULL", "PUSH-PULL"]),
    ("CRM", ["CRM", "6HI", "6 HI", "6 STAND", "COLD ROLLING"]),
    ("SPM", ["SPM", "SKIN PASS"]),
    ("FINISHING", ["TRIMMING", "REWINDING", "SLITTING", "CTL", "CUT TO LENGTH", "TENSION LEVEL"]),
]
FAMILY_NAMES = [f for f, _ in FAMILIES] + ["OTHER"]


def family_of(text) -> str:
    t = str(text or "").upper()
    for fam, kws in FAMILIES:
        if any(k in t for k in kws):
            return fam
    return "OTHER"


def to_num(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.]", "", str(v))
    try:
        return float(s) if s else np.nan
    except ValueError:
        return np.nan


# ----------------------------------------------------------------------------- week / month helpers

def month_start(d) -> pd.Timestamp:
    d = pd.Timestamp(d)
    return pd.Timestamp(d.year, d.month, 1)


def add_months(d, n: int) -> pd.Timestamp:
    d = month_start(d)
    y, m = divmod(d.month - 1 + int(n), 12)
    return pd.Timestamp(d.year + y, m + 1, 1)


def week_label(month) -> str:
    """Planner's convention: the (Sunday-start) week holding the month's last day, ISO-numbered.
    Reproduces his labels: Sep-26->WW40 2026, Dec-26->WW53 2026, Jan-27->WW5 2027, May-27->WW22 2027."""
    m = month_start(month)
    last = dt.date(m.year, m.month, calendar.monthrange(m.year, m.month)[1])
    iso = (last + dt.timedelta(days=1)).isocalendar()
    return f"WW{iso[1]} {iso[0]}"


def sales_date(month) -> dt.date:
    """Sales file stores dispatch month as the 28th."""
    m = month_start(month)
    return dt.date(m.year, m.month, 28)


def parse_week_label(s):
    m = re.match(r"\s*WW\s*(\d+)\s+(\d{4})", str(s or ""))
    if not m:
        return None
    d = dt.date.fromisocalendar(int(m[2]), int(m[1]), 1)  # Monday of that ISO week
    # the label week contains the month's last day -> month of (Monday - 1 day) is the dispatch month
    d2 = d - dt.timedelta(days=1)
    return pd.Timestamp(d2.year, d2.month, 1)


# ----------------------------------------------------------------------------- nice rounding

def nice_round(v: float) -> float:
    if not v or v <= 0:
        return 0.0
    if v < 1e6:        # < 10 lakh -> nearest 50k (min 50k)
        step = 5e4
    elif v < 1e7:      # < 1 Cr -> nearest 1 lakh
        step = 1e5
    elif v < 1e8:      # < 10 Cr -> nearest 5 lakh
        step = 5e5
    else:
        step = 1e6
    return max(step, round(v / step) * step)


# ----------------------------------------------------------------------------- knowledge base

@dataclass
class Knowledge:
    hist: pd.DataFrame            # one row per billed assembly in sales file
    labels: pd.DataFrame          # LoE rows from reference plans with billed flag
    sales_path: str
    loaded_at: dt.datetime = field(default_factory=dt.datetime.now)
    pi: pd.DataFrame = field(default_factory=pd.DataFrame)   # priority index (category & lead time)

    def summary(self) -> dict:
        h = self.hist
        return {
            "projects": int(h.project.nunique()),
            "assembly_rows": int(len(h)),
            "reference_plans": int(self.labels.project.nunique()) if len(self.labels) else 0,
            "labelled_rows": int(len(self.labels)),
        }


def load_sales(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="PROJECTS")
    df.columns = [re.sub(r"\s+", " ", str(c)).strip().upper() for c in df.columns]

    def col(*cands):
        for c in df.columns:
            if any(c.startswith(x) for x in cands):
                return c
        return None

    c_bill = next((c for c in df.columns if c.startswith("BILLING") and "$" not in c), None)
    c_usd = next((c for c in df.columns if c.startswith("BILLING") and "$" in c), None)

    def txt(c):
        if not c:
            return ""
        return df[c].map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()) \
                    .replace({"nan": "", "None": "", "NaT": "", "<NA>": ""})

    out = pd.DataFrame({
        "project": df[col("PROJECT")].astype(str).str.strip(),
        "boi": txt(col("BOI")),
        "pm": txt(col("PM")),
        "name": df[col("ASSEMBLY")].astype(str).str.strip(),
        "risk": txt(col("RISK")),
        "bill": pd.to_numeric(df[c_bill], errors="coerce"),
        "bill_usd": pd.to_numeric(df[c_usd], errors="coerce") if c_usd else np.nan,
        "status": df[col("DISP")].astype(str).str.strip(),
        "month": pd.to_datetime(df[col("DISPATCH MONTH")], errors="coerce"),
        "remarks": txt(col("REMARKS")),
        "vendor": txt(col("VENDOR")),
        "weight": pd.to_numeric(df[col("WEIGHT")], errors="coerce") if col("WEIGHT") else np.nan,
    })
    out["dispatched"] = out.status.str.upper().str.replace(".", "", regex=False).str.strip().eq("DISP")
    out = out[(out.project != "nan") & (out.name != "nan")].copy()
    out["pno"] = out.project.str.extract(r"^(\w+)")[0]
    out["line"] = out.project.str.split("-", n=2).str[-1].str.strip()
    out["family"] = out.project.map(family_of)
    out["spare"] = out.project.str.upper().str.contains("SPARE")
    out["key"] = out.name.map(norm)
    out["month"] = out.month.dt.to_period("M").dt.to_timestamp()
    return out


def build_hist(sales: pd.DataFrame) -> pd.DataFrame:
    h = sales[(sales.bill > 0) & ~sales.spare].copy()
    tot = h.groupby("project").bill.transform("sum")
    h["share"] = h.bill / tot
    # relative timing inside its project (value-weighted 5th..95th percentile window)
    pos = pd.Series(np.nan, index=h.index)
    for p, g in h.dropna(subset=["month"]).groupby("project"):
        mi = g.month.dt.year * 12 + g.month.dt.month
        lo, hi = np.percentile(mi, 5), np.percentile(mi, 95)
        pos.loc[g.index] = 0.5 if hi <= lo else ((mi - lo) / (hi - lo)).clip(0, 1)
    h["pos"] = pos
    h["is_boi"] = h.boi.str.upper().str.startswith("BOI")
    return h.reset_index(drop=True)


# ----------------------------------------------------------------------------- LoE parsing

_HDR_NAME = ("GA/ASSEMBLY", "ASSEMBLY NAME", "DESCRIPTION", "GA/ ASSEMBLY")
_BOI_MARK = ("BO SPECIFICATION", "BOI SPECIFICATION", "BOUGHT OUT")
_STOP_MARK = ("REV NO.", "REVISION NOTE", "DOCUMENTS REQUIRED", "OFFER WE RECIVED", "OFFER WE RECEIVED")
_DOC_WORDS = re.compile(r"\b(?:LAYOUT|LOAD DATA|P&ID|PID|UTILITY DATA|ROLL LIST|ASSIGNMENT|MCL|CABLE SCHEDULE|"
                        r"SPECIFICATION|CYLINDER LIST|MOTOR LIST|VALVE LIST|INSTRUMENT LIST|DATASHEET|DATA SHEET)\b")


def _cell(v):
    return re.sub(r"\s+", " ", str(v).replace("_x000d_", " ")).strip() if v is not None else ""


@dataclass
class LoE:
    rows: pd.DataFrame
    info: dict


def parse_loe(path_or_buf, sheet=None) -> LoE:
    wb = openpyxl.load_workbook(path_or_buf, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    raw = [list(r) for r in ws.iter_rows(values_only=True)]

    info = {}
    for r in raw[:25]:
        cells = [_cell(c) for c in r]
        nonempty = [c for c in cells if c]
        if len(nonempty) >= 2:
            k = nonempty[0].upper()
            for key in ("PROJECT NUMBER", "CLIENT", "LINE TYPE", "STRIP WIDTH", "LINE SPEED", "TPH"):
                if k.startswith(key) and key not in info:
                    info[key] = nonempty[1]

    hi = None
    for i, r in enumerate(raw[:60]):
        up = [_cell(c).upper() for c in r]
        if any(any(h in c for h in _HDR_NAME) for c in up) and any(c.startswith(("SR", "S.NO", "S NO", "S.N")) for c in up):
            hi = i
            break
    if hi is None:
        raise ValueError("Could not find the LoE header row (needs 'SR. NO.' and 'GA/ASSEMBLY NAME' / 'DESCRIPTION').")
    hdr = [_cell(c).upper() for c in raw[hi]]
    sub = [_cell(c).upper() for c in raw[hi + 1]] if hi + 1 < len(raw) else []

    def find(pred, start=0):
        for j, c in enumerate(hdr):
            if j >= start and pred(c):
                return j
        return None

    c_sr = find(lambda c: c.startswith(("SR", "S.NO", "S NO", "S.N")))
    c_name = find(lambda c: any(h in c for h in _HDR_NAME))
    c_sec = find(lambda c: "SECTION" in c)
    if c_sec is None:
        c_sec = find(lambda c: c.startswith("EQUIPMENT") and (not sub or "WT" not in (sub[hdr.index(c)] if hdr.index(c) < len(sub) else "")))
        if c_sec is not None and c_sec > c_name:
            c_sec = None
    c_qty = find(lambda c: c.startswith("QTY"))
    c_scope = find(lambda c: c.startswith("SCOPE"))
    c_rem = find(lambda c: c.startswith("REMARK"))
    c_drg = find(lambda c: "DRAWING" in c or "DOC/DRG NUMBER" in c or "PART N" in c)
    c_wt = None
    for j, c in enumerate(sub):
        if "TOTAL WT" in c:
            c_wt = j
    if c_wt is None:
        c_wt = find(lambda c: c.startswith("WEIGHT"))

    start = hi + 1
    if sub and any(("REV" in c or "WT" in c) for c in sub):
        start += 1

    recs, in_boi, section_carry = [], False, ""
    for i in range(start, len(raw)):
        r = raw[i] + [None] * 40
        cells = [_cell(c) for c in r]
        joined = " ".join(c for c in cells if c).upper()
        if not joined:
            continue
        if any(joined.startswith(m) for m in _STOP_MARK):
            if in_boi or i > start + 3:
                break
        if any(joined.startswith(m) for m in _BOI_MARK):
            in_boi = True
            continue
        name = cells[c_name] if c_name is not None else ""
        if name.upper() in ("TITLE", "DESCRIPTION", "GA/ASSEMBLY NAME") or cells[c_sr].upper().startswith(("SR", "S.NO")):
            if name.upper() == "TITLE":
                in_boi = True
            continue
        if not name:
            continue
        sec = cells[c_sec] if c_sec is not None else ""
        recs.append({
            "row": i + 1,
            "sr": cells[c_sr] if c_sr is not None else "",
            "section": sec,
            "name": name,
            "drawing": cells[c_drg] if c_drg is not None else "",
            "qty": to_num(r[c_qty]) if c_qty is not None else np.nan,
            "qty_txt": cells[c_qty] if c_qty is not None else "",
            "remark": cells[c_rem] if c_rem is not None else "",
            "scope": cells[c_scope] if c_scope is not None else "",
            "weight": to_num(r[c_wt]) if c_wt is not None else np.nan,
            "boi_list": in_boi,
        })
    df = pd.DataFrame(recs)
    if df.empty:
        raise ValueError("No assembly rows found under the header.")
    # hierarchy: '3' parent of '3.1', '3.2'
    srs = df.sr.astype(str).str.replace(r"\.0$", "", regex=True)
    df["sr"] = srs
    df["level"] = np.where(srs.str.fullmatch(r"\d+"), 0, np.where(srs.str.fullmatch(r"\d+\.\d+"), 1, 0))
    parent = []
    last_top = None
    for s, lvl, boi in zip(srs, df.level, df.boi_list):
        if lvl == 0:
            last_top = (s, boi)
            parent.append("")
        else:
            parent.append(s.split(".")[0] if last_top and last_top[0] == s.split(".")[0] and last_top[1] == boi else "")
    df["parent"] = parent
    child_parents = set(p for p in parent if p)
    df["has_children"] = [(lvl == 0 and s in child_parents and not boi) for s, lvl, boi in zip(srs, df.level, df.boi_list)]
    df["key"] = df.name.map(norm)
    return LoE(df.reset_index(drop=True), info)


def load_reference_plans(folder: str, exclude: set[str] | None = None) -> pd.DataFrame:
    """Planner's completed LoEs: rows with 'Billing Value' filled = billed."""
    out = []
    for f in glob.glob(os.path.join(folder, "*planning*.xlsx")):
        pno = re.match(r"(\d+)", os.path.basename(f))
        pno = pno[1] if pno else os.path.basename(f)
        if exclude and pno in exclude:
            continue
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
            ws = wb.worksheets[0]
            raw = [list(r) for r in ws.iter_rows(values_only=True)]
            hi = next(i for i, r in enumerate(raw[:60]) if any("BILLING" in _cell(c).upper() for c in r))
            hdr = [_cell(c).upper() for c in raw[hi]]
            c_bill = next(j for j, c in enumerate(hdr) if "BILLING" in c)
            c_week = next((j for j, c in enumerate(hdr) if "DISPATCH" in c), None)
            loe = parse_loe(f)
        except Exception:
            continue
        rows = loe.rows.copy()
        rows["bill"] = [to_num((raw[r - 1] + [None] * 40)[c_bill]) for r in rows.row]
        rows["week"] = [(raw[r - 1] + [None] * 40)[c_week] if c_week is not None else None for r in rows.row]
        rows["month"] = rows.week.map(parse_week_label)
        if rows.bill.fillna(0).gt(0).sum() == 0:
            continue  # plan not filled yet
        rows["billed"] = rows.bill.fillna(0) > 0
        rows["share"] = rows.bill.fillna(0) / rows.bill.fillna(0).sum()
        mi = rows.month.dt.year * 12 + rows.month.dt.month
        lo, hi_ = mi.min(), mi.max()
        rows["pos"] = ((mi - lo) / (hi_ - lo)).clip(0, 1) if hi_ > lo else 0.5
        rows["project"] = pno
        rows["family"] = family_of(" ".join(str(v) for v in loe.info.values()) or pno)
        out.append(rows)
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame(
        columns=["project", "name", "key", "billed", "bill", "section", "scope", "level", "has_children", "family"])


PI_SHEET_FAMILY = {"SKIN PASS": "SPM", "TRIMMING": "FINISHING", "CRM": "CRM", "CCL": "CCL", "PICKLING": "PICKLING",
                   "CGL": "CGL", "ARP": "ARP"}
_LONG_LEAD_WORDS = re.compile(r"\b(?:GEAR ?BOX|MANDREL|CASTING|FORGING|MILL HOUSING|HOUSING|SPINDLE|CHOCK|BACK ?UP ROLL|"
                              r"WORK ROLL|ROLLS?$|UNCOILER|RECOILER|TENSION REEL|BRIDLE)\b")


def lead_months(text) -> float:
    """'2-3M' -> 3, '5-6m' -> 6, '60' (days) -> 2, 'D-15d, mfg. 4M' -> 4.5"""
    t = str(text or "").upper()
    if not t.strip():
        return np.nan
    days = sum(float(d) for d in re.findall(r"(\d+(?:\.\d+)?)\s*D\b", t))
    t2 = re.sub(r"(\d+(?:\.\d+)?)\s*D\b", " ", t)
    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", t2)]
    if not nums:
        return round(days / 30, 1) if days else np.nan
    m = max(nums)
    if "M" not in t2 and m > 12:          # plain number of days
        m = m / 30
    return round(m + days / 30, 1)


def load_priority_index(path: str) -> pd.DataFrame:
    """Planner's Priority Index: per line type, each assembly's category (0-5) and lead time."""
    if not path or not os.path.exists(path):
        return pd.DataFrame(columns=["family", "name", "key", "category", "lead", "lead_m", "boi", "remark"])
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb.worksheets:
        fam = next((f for k, f in PI_SHEET_FAMILY.items() if k in ws.title.upper()), None)
        if not fam:
            continue
        raw = [list(r) + [None] * 30 for r in ws.iter_rows(values_only=True)]
        if len(raw) < 3:
            continue
        h0 = [_cell(c).upper() for c in raw[0]]
        h1 = [_cell(c).upper() for c in raw[1]]
        c_p = next((i for i, c in enumerate(h0) if c.startswith("PRIORITY")), None)
        c_l = next((i for i, c in enumerate(h0) if c.startswith("LEAD")), None)
        c_r = next((i for i, c in enumerate(h0) if c.startswith(("REMARK", "COMMENT"))), None)
        c_b = next((i for i, c in enumerate(h0) if c.startswith("BOI")), None)
        c_d = next((i for i, c in enumerate(h1) if "DESCRIPTION" in c), None)
        if c_p is None or c_d is None:
            continue
        parent_cat, parent_lead = None, None
        for r in raw[2:]:
            sr, name = _cell(r[0]), _cell(r[c_d])
            if not name:
                continue
            cat = _cell(r[c_p])
            lead = _cell(r[c_l]) if c_l is not None else ""
            is_top = bool(re.fullmatch(r"\d+(\.0)?", sr))
            if cat not in ("0", "1", "2", "3", "4", "5"):
                cat = ""
            if is_top:
                parent_cat, parent_lead = cat or None, lead or None
            elif not cat and parent_cat:
                cat, lead = parent_cat, lead or parent_lead or ""
            boi_txt = " ".join(_cell(r[i]) for i in range(c_d + 1, min(len(r), (c_b or c_p) + 3)) if i != c_p).upper()
            out.append({"family": fam, "name": name, "key": norm(name), "category": cat, "lead": lead,
                        "lead_m": lead_months(lead), "boi": ("BOI" in boi_txt or "TRADING" in boi_txt),
                        "remark": _cell(r[c_r]) if c_r is not None else "", "top": is_top})
    df = pd.DataFrame(out)
    return df[df.category != ""].reset_index(drop=True) if len(df) else df


def load_knowledge(sales_path: str, plans_folder: str, exclude: set[str] | None = None,
                   priority_index: str | None = None) -> Knowledge:
    sales = load_sales(sales_path)
    if exclude:
        sales = sales[~sales.pno.isin(exclude)]
    if priority_index is None:
        cands = glob.glob(os.path.join(plans_folder, "Priority Index*.xlsx"))
        priority_index = cands[0] if cands else ""
    kb = Knowledge(hist=build_hist(sales), labels=load_reference_plans(plans_folder, exclude), sales_path=sales_path)
    kb.pi = load_priority_index(priority_index)
    return kb


# ----------------------------------------------------------------------------- planning

@dataclass
class POInfo:
    project: str               # e.g. "2619-AMNS-CGL"
    family: str
    po_value: float            # in PO currency, excl. tax
    currency: str = "INR"
    fx: float = 95.0
    billable_pct: float = 100.0
    po_date: dt.date = dt.date.today()
    delivery_date: dt.date | None = None
    start_month: dt.date | None = None   # override dispatch window
    end_month: dt.date | None = None
    pm: str = ""

    @property
    def target_inr(self) -> float:
        v = self.po_value * (self.fx if self.currency == "USD" else 1.0)
        return v * self.billable_pct / 100.0


# calibrated on planner's 2611 & 2618: first dispatch at ~0.75 x delivery period after PO,
# last dispatch ~2 months after the PO delivery date (long-lead bought-outs)
WINDOW_START_FRAC = 0.75
WINDOW_TAIL_MONTHS = 2
LEAD_FLOOR = True   # use Priority Index lead time as earliest dispatch


def default_window(po: POInfo) -> tuple[pd.Timestamp, pd.Timestamp]:
    po_m = month_start(po.po_date)
    dl = month_start(po.delivery_date) if po.delivery_date else add_months(po_m, 10)
    months = max(1, (dl.year - po_m.year) * 12 + dl.month - po_m.month)
    start = add_months(po_m, int(round(WINDOW_START_FRAC * months)))
    end = add_months(dl, WINDOW_TAIL_MONTHS if months >= 4 else 0)
    today = month_start(pd.Timestamp.today())
    if start < today:                 # never plan into the past
        shift = (today.year - start.year) * 12 + today.month - start.month
        start, end = add_months(start, shift), add_months(end, shift)
    return start, max(end, start)


def _match(keys: list[str], pool: pd.Series, limit=12, cutoff=60):
    """For each key, list of (pool_index, score)."""
    uniq = pool.drop_duplicates()
    choices = uniq.tolist()
    idx_by_key = pool.groupby(pool).groups
    res = []
    for k in keys:
        if not k:
            res.append([])
            continue
        hits = process.extract(k, choices, scorer=fuzz.token_set_ratio, limit=limit, score_cutoff=cutoff)
        # penalise length mismatch (token_set_ratio gives 100 for subsets)
        out = []
        for ch, sc, _ in hits:
            ratio = fuzz.ratio(k, ch)
            score = 0.5 * sc + 0.5 * ratio
            if score < cutoff:
                continue
            for ix in idx_by_key[ch]:
                out.append((ix, score))
        out.sort(key=lambda t: -t[1])
        res.append(out)
    return res


def plan(loe: LoE, po: POInfo, kb: Knowledge) -> pd.DataFrame:
    df = loe.rows.copy()
    h = kb.hist
    fam_h = h[h.family == po.family]
    lab = kb.labels

    # ---- 1. billable? --------------------------------------------------------------------------
    hard_no = (df.section.str.upper().str.contains("ENGINEERING")
               | df.name.str.upper().str.contains(_DOC_WORDS))
    lab_leaf = lab[~lab.has_children.astype(bool)] if len(lab) else lab
    lab_grp = lab[lab.has_children.astype(bool)] if len(lab) else lab
    none = [[] for _ in range(len(df))]
    leaf_hits = _match(df.key.tolist(), lab_leaf.key, limit=6, cutoff=70) if len(lab_leaf) else none
    grp_hits = _match(df.key.tolist(), lab_grp.key, limit=6, cutoff=70) if len(lab_grp) else none
    hist_hits = _match(df.key.tolist(), fam_h.key if len(fam_h) else h.key, limit=8, cutoff=60)

    def label_vote(hits, pool):
        near = [(ix, s) for ix, s in hits if s >= 85]
        if not near:
            return None
        return any(pool.at[ix, "billed"] for ix, _ in near)   # billed anywhere -> billable

    sectioned = (df.section != "").mean() > 0.5
    new_family = len(fam_h) < 50          # no real history for this line type
    # scope-matrix LoE (revamps like 2616): every row says who supplies it
    scope_matrix = df.scope.str.upper().str.contains("SUPPLY").sum() >= 3
    billed, why = [], []
    for i, r in df.iterrows():
        hp = hist_hits[i][0][1] if hist_hits[i] else 0
        if hard_no[i]:
            billed.append(False); why.append("document / drawing")
            continue
        if scope_matrix:
            ours = "SUPPLY" in str(r.scope).upper() or re.search(r"\bYD\b", f"{r.qty_txt} {r.remark}".upper())
            billed.append(bool(ours))
            why.append("supply in our scope" if ours else "not in our supply scope")
            continue
        if r.has_children:
            billed.append(None); why.append("group")  # decided after its children
            continue
        v = label_vote(leaf_hits[i], lab_leaf)
        if v is True:
            billed.append(True); why.append("valued in earlier plan")
        elif "SUPPLY" in str(r.scope).upper():
            billed.append(True); why.append("supply in our scope")
        elif new_family and not str(r.scope).upper().startswith("CLIENT"):
            billed.append(True); why.append("new line type - included by default")
        elif v is False:
            billed.append(False); why.append("not valued in earlier plans")
        elif sectioned and not r.section and r.level == 0 and not r.boi_list:
            # un-sectioned top-level rows in a sectioned LoE = detail list of an equipment valued as a whole
            billed.append(False); why.append("detail of an equipment valued as a whole")
        elif hp >= 80:
            billed.append(True); why.append("in sales history")
        elif "SUPPLY" in str(r.scope).upper():
            billed.append(True); why.append("supply in our scope")
        elif r.boi_list:
            billed.append(True); why.append("bought-out list")
        else:
            billed.append(False); why.append("no history")
    df["billable"] = billed
    df["reason"] = why
    # group headers. Planner either values the group as one unit (UNCOILER ASSY) or each child
    # (BRIDLE ROLL ASSY.#1..#10). Follow his earlier plans; otherwise follow the children.
    for i, r in df[df.billable.isna()].iterrows():
        kids = df[(df.parent == r.sr) & (df.boi_list == r.boi_list) & (df.index != i)]
        v = label_vote(grp_hits[i], lab_grp)
        hp = hist_hits[i][0][1] if hist_hits[i] else 0
        kid_bill = kids.billable.fillna(False).astype(bool)
        if v is True or (v is None and not kid_bill.any() and hp >= 70):
            df.at[i, "billable"] = True
            df.at[i, "reason"] = "group valued as one unit" + (" (as in earlier plan)" if v else "")
            for k in kids.index[kid_bill.values]:
                if df.at[k, "reason"] != "valued in earlier plan":   # keep only items he values separately
                    df.at[k, "billable"] = False; df.at[k, "reason"] = "part of group"
        elif kid_bill.any():
            df.at[i, "billable"] = False; df.at[i, "reason"] = "group - valued via sub-assemblies"
        else:
            df.at[i, "billable"] = False; df.at[i, "reason"] = "group - no history"
    df["billable"] = df.billable.astype(bool)

    # differently-named rows in one section that all map to the same past item (CAG Structure -
    # Columns / Beams / Staircases -> 'CAG STRUCTURE'): value only the best one, like the planner does.
    hpool = fam_h if len(fam_h) else h
    df["_hkey"] = [hpool.at[hh[0][0], "key"] if hh else "" for hh in hist_hits]
    df["_hsc"] = [hh[0][1] if hh else 0 for hh in hist_hits]
    cand = df[df.billable & (df._hkey != "") & (df.key != df._hkey)]
    for (_, _), g in cand.groupby(["section", "_hkey"]):
        if g.key.nunique() < 2:
            continue
        from_plan = g.reason.str.contains("earlier plan")
        keep = g[from_plan].index if from_plan.any() else [g._hsc.idxmax()]
        for k in g.index:
            if k not in keep and df.at[k, "reason"] == "in sales history":
                df.at[k, "billable"] = False; df.at[k, "reason"] = f"covered by '{df.loc[keep[0], 'name'][:30]}'"
    df = df.drop(columns=["_hkey", "_hsc"])

    # ---- 2. value & timing from history --------------------------------------------------------
    pool = fam_h if len(fam_h) >= 50 else h
    hits_all = _match(df.key.tolist(), pool.key, limit=10, cutoff=55)
    hits_any = _match(df.key.tolist(), h.key, limit=10, cutoff=70) if pool is not h else hits_all
    shares, poss, boi, basis, conf, absv = [], [], [], [], [], []
    lab_b = lab[lab.billed.astype(bool)] if len(lab) else lab
    lab_b_hits = _match(df.key.tolist(), lab_b.key, limit=6, cutoff=85) if len(lab_b) else none
    for i, r in df.iterrows():
        # planner's own earlier plans first (same kind of row: group vs leaf)
        lh = [(ix, s) for ix, s in lab_b_hits[i] if bool(lab_b.at[ix, "has_children"]) == bool(r.has_children)]
        if lh:
            g = lab_b.loc[[ix for ix, _ in lh]]
            hist_ok = [(ix, s) for ix, s in hits_all[i] if s >= 85]
            share = float(np.median(g.share))
            pos_ = float(np.nanmedian(g.pos)) if g.pos.notna().any() else np.nan
            shares.append(share); poss.append(pos_); absv.append(float(np.median(g.bill)))
            boi.append(bool(r.boi_list) or (bool(pool.loc[[ix for ix, _ in hist_ok]].is_boi.mean() >= .5) if hist_ok else False))
            basis.append(f"planner's plan {', '.join(sorted(set(g.project)))}: '{g.name.iloc[0][:40]}'")
            conf.append(round(lh[0][1]))
            continue
        strong = [(ix, s) for ix, s in hits_all[i] if s >= 70]
        if strong:
            hits, use_pool = strong, pool          # same line family
        else:
            hits, use_pool = hits_any[i], h        # any line type
        if not hits:
            shares.append(np.nan); poss.append(np.nan); boi.append(r.boi_list); basis.append(""); conf.append(0)
            absv.append(np.nan)
            continue
        top = hits[0][1]
        sel = [(ix, s) for ix, s in hits if s >= top - 8][:12]
        g = use_pool.loc[[ix for ix, _ in sel]]
        absv.append(float(np.median(g.bill)))
        shares.append(float(np.median(g.share)))
        poss.append(float(np.nanmedian(g.pos)) if g.pos.notna().any() else np.nan)
        boi.append(bool(g.is_boi.mean() >= 0.5) or r.boi_list)
        projs = sorted(set(g.pno))
        basis.append(f"{len(g)} past rows ({', '.join(projs[:4])}{'…' if len(projs) > 4 else ''}) e.g. '{g.name.iloc[0][:40]}'")
        conf.append(round(top))
    df["hist_share"] = shares
    df["hist_pos"] = poss
    loe_boi = (df.remark.str.upper().str.contains(r"\bBOI\b|TRADING", regex=True)
               | df.scope.str.upper().str.contains(r"\bBOI\b", regex=True)
               | df.qty_txt.str.upper().str.contains(r"\bBOI\b", regex=True))
    df["is_boi"] = pd.Series(boi, index=df.index).astype(bool) | loe_boi
    df["basis"] = basis
    df["match_score"] = conf
    # long-lead bought-out share among matched history rows (BOI-LL)
    df["hist_ll"] = [
        float((pool.loc[[ix for ix, s in hits_all[i] if s >= 80]].boi.str.upper() == "BOI-LL").mean())
        if any(s >= 80 for _, s in hits_all[i]) else 0.0 for i in df.index]

    # ---- 2b. category (0,1,2,3,4,5) & lead time ------------------------------------------------
    pi = getattr(kb, "pi", pd.DataFrame())
    pi_f = pi[pi.family == po.family] if len(pi) else pi
    pi_hits = _match(df.key.tolist(), pi_f.key, limit=3, cutoff=85) if len(pi_f) else none
    cats, leads, cat_basis = [], [], []
    by_sr = {}
    for i, r in df.iterrows():
        cat, lead, cb = None, np.nan, ""
        if pi_hits[i]:
            p = pi_f.loc[pi_hits[i][0][0]]
            cat, lead, cb = p.category, p.lead_m, f"Priority Index ({po.family}): '{p['name'][:35]}'"
        elif r.parent and r.parent in by_sr and by_sr[r.parent][2].startswith("Priority"):
            cat, lead, cb = by_sr[r.parent][0], by_sr[r.parent][1], "inherited from parent assembly"
        if cat is None:
            u = r["name"].upper()
            if r.is_boi:
                long_ = r.hist_ll >= 0.5 or bool(re.search(r"POT EQUIP|POT ROLL|INDUCTION|SKIN PASS|TENSION LEVEL|"
                                                            r"WELDER|AIR KNIFE|HYDRAULIC|ELECTRICAL|PREMELT|MAIN POT", u))
                cat, cb = ("5", "rule: long-lead bought-out") if long_ else ("3/4", "rule: bought-out")
            elif _LONG_LEAD_WORDS.search(u):
                cat, cb = "2", "rule: manufactured, long-lead content"
            else:
                cat, cb = "1", "rule: manufactured, short lead"
        by_sr[r.sr] = (cat, lead, cb)
        cats.append(cat); leads.append(lead); cat_basis.append(cb)
    df["category"] = cats
    df["lead_m"] = leads
    df["category_basis"] = cat_basis

    # ---- 3. scale to PO value ------------------------------------------------------------------
    # Whole lines: split the PO by each assembly's typical share of a project.
    # Small orders (revamps, modifications, spares): shares are meaningless -> use typical item
    # prices from history, then scale to the PO.
    b = df.billable
    target = po.target_inr
    df["hist_abs"] = absv
    fam_tot = (fam_h if len(fam_h) else h).groupby("project").bill.sum()
    small_order = len(fam_tot) == 0 or target < 0.2 * float(fam_tot.median())
    df.attrs["value_mode"] = "item prices" if small_order else "project shares"
    if small_order:
        ref = fam_h if len(fam_h) else h
        typ = {
            "LL": ref.loc[ref.boi.str.upper() == "BOI-LL", "bill"].median(),
            "BOI": ref.loc[ref.boi.str.upper() == "BOI", "bill"].median(),
            "MANF": ref.loc[~ref.boi.str.upper().str.startswith("BOI"), "bill"].median(),
        }
        typ = {k: (v if pd.notna(v) else h.bill.median()) for k, v in typ.items()}
        default = np.where(df.category == "5", typ["LL"], np.where(df.is_boi, typ["BOI"], typ["MANF"]))
        raw_v = df.hist_abs.where(df.hist_abs.notna(), pd.Series(default, index=df.index)).where(b, 0.0)
    else:
        fill_share = np.nanmedian(df.loc[b, "hist_share"]) if df.loc[b, "hist_share"].notna().any() else 0.01
        sh = df.hist_share.where(df.hist_share.notna(), fill_share)
        raw_v = sh.where(b, 0.0)
    est = raw_v / raw_v.sum() * target if raw_v.sum() > 0 else raw_v
    vals = est.map(nice_round).where(b, 0.0)
    diff = target - vals.sum()
    if b.any() and abs(diff) > 0:
        j = vals.idxmax()
        vals.at[j] = max(0.0, vals.at[j] + diff)
    df["billing_value"] = vals

    # ---- 4. dispatch month ---------------------------------------------------------------------
    start, end = default_window(po)
    if po.start_month:
        start = month_start(po.start_month)
    if po.end_month:
        end = month_start(po.end_month)
    span = (end.year - start.year) * 12 + end.month - start.month
    pos = df.hist_pos.fillna(np.nanmedian(df.hist_pos) if df.hist_pos.notna().any() else 0.5)
    # bought-out without history -> late half
    pos = pos.where(~(df.hist_pos.isna() & df.is_boi), 0.75)
    months = [add_months(start, int(round(p * span))) if bb else pd.NaT for p, bb in zip(pos, b)]
    if LEAD_FLOOR:
        # never earlier than PO + design (1 month) + its lead time, capped at the window end
        po_m = month_start(po.po_date)
        for k, (m, lm, bb) in enumerate(zip(months, df.lead_m, b)):
            if bb and pd.notna(lm):
                floor = add_months(po_m, int(np.ceil(lm)) + 1)
                if floor > m:
                    months[k] = min(floor, end)
    df["dispatch_month"] = months
    df["dispatch_week"] = [week_label(m) if bb else "" for m, bb in zip(months, b)]
    df["type"] = np.where(df.is_boi, "BOI", "MANF.")
    return df


def plan_lots(po: POInfo, n_lots: int, label: str = "DISPATCH LOT") -> pd.DataFrame:
    """Lot-wise plan (how the planner books export / trading / revamp orders, e.g. 2636: 3 equal lots)."""
    start, end = default_window(po)
    if po.start_month:
        start = month_start(po.start_month)
    if po.end_month:
        end = month_start(po.end_month)
    span = (end.year - start.year) * 12 + end.month - start.month
    n = max(1, int(n_lots))
    target = po.target_inr
    each = nice_round(target / n)
    vals = [each] * n
    vals[-1] = target - each * (n - 1)
    months = [add_months(start, int(round(k * span / (n - 1)))) if n > 1 else start for k in range(n)]
    return pd.DataFrame({
        "row": range(1, n + 1), "sr": [str(k + 1) for k in range(n)], "section": "", "name": [f"{label}-{k + 1}" for k in range(n)],
        "qty": 1, "scope": "", "billable": True, "billing_value": vals, "dispatch_month": months,
        "dispatch_week": [week_label(m) for m in months], "type": "BOI", "reason": "lot-wise plan",
        "basis": f"{n} equal lots", "match_score": 0, "boi_list": False, "level": 0, "has_children": False,
    })


def monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    d = df[df.billable & (df.billing_value > 0)]
    if d.empty:
        return pd.DataFrame(columns=["Month", "Week", "Assemblies", "Billing (₹ Cr)"])
    g = d.groupby("dispatch_month").agg(n=("name", "size"), v=("billing_value", "sum")).reset_index()
    return pd.DataFrame({
        "Month": g.dispatch_month.dt.strftime("%b %Y"),
        "Week": g.dispatch_month.map(week_label),
        "Assemblies": g.n,
        "Billing (₹ Cr)": (g.v / 1e7).round(2),
    })


INDIAN_FMT = '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0'


def write_plan_into_loe(loe_bytes: bytes, sheet: str | None, df: pd.DataFrame) -> bytes:
    """Original LoE with 'Dispatch Week' and 'Billing Value (INR)' filled in - the planner's own format.
    Saved as '<project> planning.xlsx' next to the sales file it becomes a reference plan."""
    import io
    wb = openpyxl.load_workbook(io.BytesIO(loe_bytes))
    ws = wb[sheet] if sheet else wb.worksheets[0]
    hdr_row = None
    for r in range(1, 61):
        vals = [_cell(ws.cell(r, c).value).upper() for c in range(1, ws.max_column + 1)]
        if any(any(h in v for h in _HDR_NAME) for v in vals) and any(v.startswith(("SR", "S.NO", "S NO", "S.N")) for v in vals):
            hdr_row = r
            break
    if hdr_row is None:
        raise ValueError("LoE header row not found")
    cols = {_cell(ws.cell(hdr_row, c).value).upper(): c for c in range(1, ws.max_column + 1)}

    def col_for(name):
        for k, c in cols.items():
            if k.startswith(name.upper()):
                return c
        c = ws.max_column + 1
        ws.cell(hdr_row, c, name)
        cols[name.upper()] = c
        return c

    c_w, c_b = col_for("Dispatch Week"), col_for("Billing Value (INR)")
    for _, r in df.iterrows():
        if pd.isna(r.get("row")):
            continue
        valued = bool(r.billable) and r.billing_value > 0
        ws.cell(int(r.row), c_w, r.dispatch_week if valued else None)
        ws.cell(int(r.row), c_b, float(r.billing_value) if valued else None)
        if valued:
            ws.cell(int(r.row), c_b).number_format = INDIAN_FMT
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def sales_rows(df: pd.DataFrame, po: POInfo) -> pd.DataFrame:
    """Rows in the PROJECTS-sheet layout of SALES & CASH FLOW.xlsx, ready to paste."""
    d = df[df.billable & (df.billing_value > 0)]
    return pd.DataFrame({
        "PROJECT": po.project,
        "BOI /MANF.": np.where(d.type == "BOI", "BOI", ""),
        "PM": po.pm,
        "ASSEMBLY": d.name,
        "RISK FB/ PAYMENT/ RED FLAG": "",
        "BILLING VALUE ($)": "",
        "BILLING VALUE": d.billing_value.round(0),
        "DISP. STATUS": "Not Disp.",
        "DISPATCH MONTH": d.dispatch_month.map(sales_date),
        "REMARKS": "",
    })
