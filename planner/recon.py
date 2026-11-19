"""Forecast vs actual: match every project's SALES & CASH FLOW rows (the forecast) to SAP billing (the actual).

SAP (`revenue data sap.xlsx`, sheet Query1) is read-only and reliable. One WBS element = one assembly; it can carry
several invoice lines (part bills, cancellations S1, credit memos G2, re-bills), so SAP is netted per WBS first.
The sales file names assemblies differently ("COIL CAR ASSY. @ UNCOILER" vs "UNCOILER COIL CAR ASSEMBLY"), so rows
are matched on a cleaned name, helped by value and month. SAP often bills several sales rows under one WBS (and vice
versa), so values that add up are matched too. Matched items form a group; each group gets a status.

Manual fixes (link / unlink / note) are kept in recon_overrides.json and applied on top of the automatic match.
"""
from __future__ import annotations

import datetime as dt
import itertools
import json
import os
import re
import shutil
import tempfile
import threading
from collections import Counter

import numpy as np
import pandas as pd
from rapidfuzz import fuzz
from rapidfuzz.process import cdist

import engine as E

HERE = os.path.dirname(os.path.abspath(__file__))
OVERRIDES = os.path.join(HERE, "recon_overrides.json")
_olock = threading.Lock()

VALUE_TOL = 0.05          # values within 5% (or ₹50k) count as equal
VALUE_TOL_ABS = 50_000
PART_BILLED = 0.90        # SAP net below 90% of the forecast = partly billed

# ----------------------------------------------------------------------------- names

_ABBR = [(re.compile(a), b) for a, b in [
    (r"\b(?:DELIVERY|EXIT)\s+TENSION\s+REEL\b|\bTENSION\s+REEL\b[^()]*\(\s*(?:DELIVERY|EXIT)\s*\)", " DTR "),
    (r"\bENTRY\s+TENSION\s+REEL\b|\bTENSION\s+REEL\b[^()]*\(\s*ENTRY\s*\)", " ETR "),
    (r"\bPAY\s*-?\s*OFF\s+REEL\b", " POR "),
    (r"\bACC\b\.?", " ACCUMULATOR "),
    (r"\bDEF\b\.?", " DEFLECTOR "),
    (r"\bDELIVERY\b", " EXIT "),
    (r"\bPLATE\s?FORM\b", " PLATFORM "),
    (r"\bSEGE?MENT\b", " "),
    (r"\bLOT\b", " "),
    (r"\bEXHU?ST\b", " EXHAUST "),
    (r"\bKINFE\b", " KNIFE "),
    (r"\bM/C\b", " MACHINE "),
    (r"\bCOUNTER\s*WEIGHT\b", " COUNTERWEIGHT "),
    (r"\bBACK\s*-?\s*UP\b", " BACKUP "),
    (r"\bPAY\s*-?\s*OFF\b", " PAYOFF "),
    (r"\bTHK\b\.?", " THICKNESS "),
    (r"\bHYD\b\.?", " HYDRAULIC "),
    (r"\bELEC\b\.?", " ELECTRICAL "),
    (r"\bOPR\b\.?", " OPERATOR "),
    (r"\bDRV\b\.?", " DRIVE "),
    (r"\bPIT\s*COVER\b", " PITCOVER "),
    (r"\bBUP\b", " BACKUP "),
    (r"\bWR\b", " WORK ROLL "),
    (r"\bIMR\b", " INTERMEDIATE ROLL "),
    (r"\bHPP\b", " HYDRAULIC POWER PACK "),
    (r"\bECR\b", " ELECTRICAL CONTROL ROOM "),
]]
_DROP = re.compile(r"ASS(?:E|EN|)M?BL(?:Y|IES)|ASSY|ASEMBLY")
_STOP = {"OF", "THE", "WITH", "AND", "FOR", "SET", "NOS", "NO", "TO", "IN", "AT", "A", "SUPPLY", "COMPLETE",
         "ITEM", "ITEMS"}
# words that must agree when both sides have one (an uncoiler is never a recoiler)
_EXCLUSIVE = [{"UNCOILER", "RECOILER"}, {"ENTRY", "EXIT", "CENTRE", "CENTER"}, {"TOP", "BOTTOM"},
              {"POR", "ETR", "DTR", "PAYOFF"}, {"FIXED", "MOVING"}, {"UPPER", "LOWER"},
              {"HORIZONTAL", "VERTICAL"}]
_GENERIC = {"ROLL", "UNIT", "SYSTEM", "FRAME", "STRUCTURE", "TABLE", "PLATE", "MAIN", "LINE", "TYPE", "DRIVE",
            "SIDE", "ENTRY", "EXIT", "TOP", "BOTTOM", "TANK", "MACHINE"}
_SERVICE = re.compile(r"ENGINEERING SERVICE|SUPPLY OF ENGINEERING|ERECTION|COMMISSIONING|SERVICE CHARGE|"
                      r"SUPERVISION|FREIGHT|JOB WORK|TRAINING|INSTALLATION", re.I)


def clean(name, pno: str = "") -> str:
    s = str(name or "").upper()
    s = re.sub(r"^\s*\d{4}\s*[-:]?\s*", " ", s)                 # 2428-BRIDLE … / 2522 BRIDLE …
    if pno:
        s = s.replace(f"({pno}", "(")
    s = re.sub(r"#\s*0*(\d+)", r" #\1 ", s)                      # #01 -> #1
    s = re.sub(r"\b(?:NO|SR)\.?\s*0*(\d+)\b", r" #\1 ", s)
    s = re.sub(r"[-_]\s*0*(\d{1,2})\s*$", r" #\1 ", s)            # SEGMENT-01 -> #1
    for rx, to in _ABBR:
        s = rx.sub(to, s)
    s = _DROP.sub(" ", s)
    s = re.sub(r"[^A-Z0-9# ]+", " ", s)
    words = [w[:-1] if len(w) > 4 and w.endswith("S") and not w.endswith("SS") else w
             for w in s.split() if w not in _STOP]
    return " ".join(words)


def _nums(s: str) -> set:
    return set(re.findall(r"#(\d+)", s))


def _toks(key: str) -> set:
    return {w for w in key.split() if len(w) >= 3 and not w.startswith("#") and w not in _GENERIC}


class _K:
    """A cleaned name, pre-split for fast scoring."""
    __slots__ = ("nums", "words", "bare")

    def __init__(self, key: str):
        self.nums = _nums(key)
        self.words = set(key.split())
        self.bare = re.sub(r"#\d+", "", key).strip()


def _penalty(a: _K, b: _K) -> float | None:
    """None = can never match (#2 vs #3, uncoiler vs recoiler, …); else points to subtract."""
    if not a.bare or not b.bare:
        return None
    if a.nums and b.nums and a.nums != b.nums:
        return None
    for grp in _EXCLUSIVE:
        ga, gb = a.words & grp, b.words & grp
        if ga and gb and not (ga & gb):
            return None
    return 4.0 if (a.nums or b.nums) and not (a.nums and b.nums) else 0.0


def _score_matrix(skeys: list[str], wkeys: list[list[str]]) -> np.ndarray:
    """Name score (0-100) of every sales key against every WBS (best of its names).
    Score = 0.45 token-set + 0.55 token-sort, so one shared word ("DTR") can't carry a match on its own."""
    flat, owner = [], []
    for j, ks in enumerate(wkeys):
        for k in ks:
            flat.append(k); owner.append(j)
    out = np.zeros((len(skeys), len(wkeys)))
    if not skeys or not flat:
        return out
    SA, SB = [_K(k) for k in skeys], [_K(k) for k in flat]
    a_bare, b_bare = [x.bare for x in SA], [x.bare for x in SB]
    m = 0.45 * cdist(a_bare, b_bare, scorer=fuzz.token_set_ratio, workers=-1) + \
        0.55 * cdist(a_bare, b_bare, scorer=fuzz.token_sort_ratio, workers=-1)
    for i, j in zip(*np.nonzero(m >= 35)):
        p = _penalty(SA[i], SB[j])
        v = 0.0 if p is None else m[i, j] - p
        if v > out[i, owner[j]]:
            out[i, owner[j]] = v
    return out


def name_score(a: str, b: str) -> float:
    """0-100 similarity of two cleaned names (see _score_matrix)."""
    return float(_score_matrix([a], [[b]])[0, 0])


# ----------------------------------------------------------------------------- SAP

LAST_GOOD = os.path.join(HERE, "cache", "sap_last.xlsx")


def _readable(path: str) -> str:
    """The SAP export is often open in Excel, which locks it even against Python's copy. Try: the file itself →
    a Windows-level copy (works while Excel has it open) → the last copy that was read successfully."""
    os.makedirs(os.path.dirname(LAST_GOOD), exist_ok=True)
    try:
        with open(path, "rb"):
            pass
        shutil.copyfile(path, LAST_GOOD)
        return path
    except PermissionError:
        pass
    tmp = os.path.join(tempfile.gettempdir(), "yd_sap_copy.xlsx")
    try:
        import subprocess
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        f"Copy-Item -LiteralPath '{path}' -Destination '{tmp}' -Force"],
                       check=True, capture_output=True, timeout=60)
        shutil.copyfile(tmp, LAST_GOOD)
        return tmp
    except Exception:
        if os.path.exists(LAST_GOOD):
            return LAST_GOOD
        raise PermissionError(f"{os.path.basename(path)} is locked (open in Excel?) and no earlier copy exists. Close it and refresh.")

def load_sap(path: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Return (lines, wbs, meta). lines = invoice lines with a project; wbs = netted per (pno, wbs)."""
    raw = pd.read_pickle(path) if path.endswith(".pkl") else pd.read_excel(_readable(path), sheet_name="Query1")
    d = pd.DataFrame({
        "doc": raw["document"].astype(str),
        "type": raw["doc_type"].astype(str).str.strip(),
        "date": pd.to_datetime(raw["Billing_date"], errors="coerce"),
        "product": raw["ProductName"].astype(str).str.strip(),
        "wbs": raw["WBSElementExternalID"],
        "wbs_desc": raw["WBSDescription"],
        "project": raw["Project"],
        "pdesc": raw["ProjectDescription"],
        "customer": raw["CustomerFullName"].astype(str),
        "cur": raw["TransactionCurrency"].astype(str),
        "amount_tc": pd.to_numeric(raw["NetAmount"], errors="coerce").fillna(0.0),
        "inr": pd.to_numeric(raw["TOTAL_SALE_WITHOUT_TAX"], errors="coerce").fillna(0.0),
    })
    pno = d.project.astype(str).str.extract(r"(\d{4})")[0]
    d["pno"] = pno.fillna(d["product"].str.extract(r"^\s*(\d{4})\s*-")[0])
    meta = {
        "rows": int(len(d)), "from": d.date.min(), "to": d.date.max(),
        "unlinked_rows": int(d.pno.isna().sum()), "unlinked_inr": float(d.loc[d.pno.isna(), "inr"].sum()),
        "all_lines": d[["date", "inr"]].copy(),          # every billing line, with or without a project (FY totals)
    }
    d = d[d.pno.notna()].copy()
    d["wbs"] = d.wbs.where(d.wbs.notna(), "ITEM:" + d["product"])
    d["wbs_desc"] = d.wbs_desc.where(d.wbs_desc.notna(), d["product"])

    groups: dict = {}
    for r in d.sort_values("date").to_dict("records"):         # plain dicts: pandas per-group calls are slow
        groups.setdefault((r["pno"], r["wbs"]), []).append(r)
    rows = []
    for (p, w), g in groups.items():
        inv = [r for r in g if r["type"] == "F2"]
        dates = [r["date"] for r in (inv or g) if pd.notna(r["date"])]
        names = list(dict.fromkeys([str(g[0]["wbs_desc"]).strip()] + [r["product"] for r in g]))
        rows.append({
            "pno": p, "wbs": str(w), "desc": names[0], "names": names,
            "keys": [clean(n, p) for n in names],
            "net": float(sum(r["inr"] for r in g)), "gross": float(sum(r["inr"] for r in inv)),
            "cancelled": float(-sum(r["inr"] for r in g if r["type"] in ("S1", "G2"))),
            "net_tc": float(sum(r["amount_tc"] for r in g)), "cur": ",".join(sorted({r["cur"] for r in g})),
            # month the item counts as billed: the last invoice (a cancelled bill is followed by a re-bill)
            "first": min(dates) if dates else pd.NaT, "last": max(dates) if dates else pd.NaT,
            "n": len(g),
            "service": bool(_SERVICE.search(names[0]) and all(_SERVICE.search(n) for n in names[1:])),
            "lines": [{"date": r["date"].strftime("%d %b %Y") if pd.notna(r["date"]) else "", "doc": r["doc"],
                       "type": r["type"], "product": r["product"], "inr": float(r["inr"])} for r in g],
        })
    wbs = pd.DataFrame(rows)
    meta["projects"] = d.groupby("pno").agg(sap_desc=("pdesc", "first"), customer=("customer", "first")).to_dict("index")
    return d, wbs, meta


# ----------------------------------------------------------------------------- overrides

def load_overrides() -> dict:
    if not os.path.exists(OVERRIDES):
        return {}
    with open(OVERRIDES, encoding="utf-8") as fh:
        return json.load(fh)


def save_override(pno: str, action: str, skey: str | None, wbs: str | None, note: str | None = None) -> dict:
    """action: link | unlink | clear | note. skey = sales row key ("NAME|n"), wbs = SAP WBS element."""
    with _olock:
        o = load_overrides()
        p = o.setdefault(pno, {})
        for k, v in (("link", []), ("unlink", []), ("notes", {})):
            p.setdefault(k, v)
        pair = [skey, wbs]
        if action in ("link", "unlink", "clear"):
            p["link"] = [x for x in p["link"] if x != pair]
            p["unlink"] = [x for x in p["unlink"] if x != pair]
            if action != "clear":
                p[action].append(pair)
        elif action == "note":
            k = skey or f"SAP:{wbs}"
            if note:
                p["notes"][k] = note
            else:
                p["notes"].pop(k, None)
        with open(OVERRIDES, "w", encoding="utf-8") as fh:
            json.dump(o, fh, indent=1, ensure_ascii=False)
        return p


def reset_items(pno: str, skeys: list[str], wbss: list[str]) -> None:
    """Forget every manual link/unlink that involves these sales rows or SAP items (back to the automatic match)."""
    with _olock:
        o = load_overrides()
        p = o.get(pno)
        if not p:
            return
        sk, wb = set(skeys), set(wbss)
        for k in ("link", "unlink"):
            p[k] = [x for x in p.get(k, []) if x[0] not in sk and x[1] not in wb]
        with open(OVERRIDES, "w", encoding="utf-8") as fh:
            json.dump(o, fh, indent=1, ensure_ascii=False)


# ----------------------------------------------------------------------------- matching

def _months_between(a, b) -> int | None:
    if a is None or b is None or pd.isna(a) or pd.isna(b):
        return None
    return int((b.year - a.year) * 12 + (b.month - a.month))


def _value_fit(f: float, s: float) -> float:
    if f <= 0 or s <= 0:
        return 0.0
    return min(f, s) / max(f, s)


def _tol(v: float) -> float:
    return max(0.005 * abs(v), 2000.0)


def _subset(vals: list[tuple], target: float, tol: float, kmax: int):
    """Smallest combination (size 1..kmax) of (id, value, weight) summing to target within tol; best weight wins."""
    for k in range(1, kmax + 1):
        best = None
        for combo in itertools.combinations(vals, k):
            if abs(sum(c[1] for c in combo) - target) <= tol:
                w = sum(c[2] for c in combo)
                if best is None or w > best[0]:
                    best = (w, combo)
        if best:
            return [c[0] for c in best[1]]
    return None


def match_project(srows: pd.DataFrame, wrows: pd.DataFrame, ov: dict | None = None) -> list[tuple[set, set, dict]]:
    """Return groups (sales index set, wbs index set, how). how[(si, wi)] = (kind, score) with kind =
    manual | name | sum (values add up) | value (same value and month, names unrelated - check it)."""
    ov = ov or {}
    unlink = {tuple(x) for x in ov.get("unlink", [])}
    S = [dict(i=i, key=k, skey=sk, bill=float(b), month=m, toks=_toks(k), disp=bool(d))
         for i, k, sk, b, m, d in zip(srows.index, srows.key, srows.skey, srows.bill, srows.month, srows.dispatched)]
    W = [dict(i=i, keys=ks, wbs=wb, net=float(n), last=l, toks=set().union(*[_toks(k) for k in ks]))
         for i, ks, wb, n, l in zip(wrows.index, wrows["keys"], wrows.wbs, wrows.net, wrows["last"])]
    sidx = {x["i"]: x for x in S}
    widx = {x["i"]: x for x in W}

    M = _score_matrix([s["key"] for s in S], [w["keys"] for w in W])
    D = _score_matrix([s["key"] for s in S], [w["keys"][:1] for w in W])     # WBS description only
    nscore, pairs = {}, []
    for a, b in zip(*np.nonzero(M >= 35)):
        s, w = S[a], W[b]
        ns = float(M[a, b])
        nscore[(s["i"], w["i"])] = ns
        if (s["skey"], w["wbs"]) in unlink or ns < 45:
            continue
        # a row not marked Disp. would become a red flag: it needs the WBS itself to match, not just one product
        # billed under it ("POWER PACK FOR SQUEEZE ROLL" under the power-pack WBS is not the squeeze roll)
        if not s["disp"] and not (D[a, b] >= 70 or ns >= 85):
            continue
        vf = _value_fit(s["bill"], w["net"])
        bonus = 20 if vf >= 0.98 else 10 if vf >= 0.9 else 4 if vf >= 0.7 else -8 if 0 < vf < 0.3 else 0
        md = _months_between(s["month"], w["last"])
        if md is not None and abs(md) <= 1:
            bonus += 5
        score = ns + bonus
        ok = (ns >= 62 and score >= 72) or (ns >= 50 and vf >= 0.98 and md is not None and abs(md) <= 1)
        if ok and (s["bill"] > 0 or ns >= 85):
            pairs.append((score, ns, s["i"], w["i"]))
    pairs.sort(key=lambda x: (-x[0], -x[1]))

    ls = {x["i"]: set() for x in S}
    lw = {x["i"]: set() for x in W}
    how: dict = {}

    def link(si, wi, kind, sc):
        ls[si].add(wi); lw[wi].add(si); how[(si, wi)] = (kind, round(float(sc), 1))

    s_by = {x["skey"]: x["i"] for x in S}
    w_by = {x["wbs"]: x["i"] for x in W}
    for sk, wk in ov.get("link", []):                       # manual links first
        if sk in s_by and wk in w_by:
            link(s_by[sk], w_by[wk], "manual", 100)
    for score, ns, si, wi in pairs:                          # one-to-one, best first
        if not ls[si] and not lw[wi]:
            link(si, wi, "name", score)

    def near(m1, m2, k=3):
        d = _months_between(m1, m2)
        return d is None or abs(d) <= k

    def affinity(s, w):
        return nscore.get((s["i"], w["i"]), 0) + 10 * len(s["toks"] & w["toks"])

    def related(s, w):
        return (s["skey"], w["wbs"]) not in unlink and \
            (bool(s["toks"] & w["toks"]) or nscore.get((s["i"], w["i"]), 0) >= 40)

    # sums: a WBS billed for several sales rows, or a sales row billed under several WBS
    for _ in range(2):
        for w in sorted(W, key=lambda x: -x["net"]):
            r = w["net"] - sum(sidx[si]["bill"] for si in lw[w["i"]])
            if w["net"] <= 0 or r <= _tol(w["net"]):
                continue
            pool = sorted([(s["i"], s["bill"], affinity(s, w)) for s in S
                           if not ls[s["i"]] and s["disp"] and s["bill"] > 0 and near(s["month"], w["last"], 2)
                           and related(s, w)],
                          key=lambda x: -x[2])[:16]
            for si in _subset(pool, r, _tol(r), 3) or []:
                link(si, w["i"], "sum", nscore.get((si, w["i"]), 0))
        for s in sorted(S, key=lambda x: -x["bill"]):
            r = s["bill"] - sum(widx[wi]["net"] for wi in ls[s["i"]])
            if not s["disp"] or s["bill"] <= 0 or r <= _tol(s["bill"]):
                continue
            pool = sorted([(w["i"], w["net"], affinity(s, w)) for w in W
                           if not lw[w["i"]] and w["net"] > 0 and near(s["month"], w["last"], 2) and related(s, w)],
                          key=lambda x: -x[2])[:16]
            for wi in _subset(pool, r, _tol(r), 3) or []:
                link(s["i"], wi, "sum", nscore.get((s["i"], wi), 0))

    # leftovers join their best strong name match even if that item is already taken
    best_s, best_w = {}, {}
    for score, ns, si, wi in pairs:
        if ns >= 80:
            best_s.setdefault(si, (wi, score))
            best_w.setdefault(wi, (si, score))
    # ... but only while the other side still has value left to cover it (a WBS billed for stands #1-#3 must not
    # swallow stand #4 that is not dispatched yet)
    def w_room(wi):
        return widx[wi]["net"] - sum(sidx[x]["bill"] for x in lw[wi])

    def s_room(si):
        return sidx[si]["bill"] - sum(widx[x]["net"] for x in ls[si])

    for si, (wi, sc) in sorted(best_s.items(), key=lambda kv: -kv[1][1]):
        b = sidx[si]["bill"]
        if not ls[si] and (b <= 0 and not lw[wi] or b > 0 and w_room(wi) >= 0.9 * b):
            link(si, wi, "name", sc)
    for wi, (si, sc) in sorted(best_w.items(), key=lambda kv: -kv[1][1]):
        n = widx[wi]["net"]
        if not lw[wi] and n > 0 and s_room(si) >= 0.9 * n:
            link(si, wi, "name", sc)

    # still unmatched on both sides and each other's best name match ("DTR ASSEMBLY WITH PUSHER PLATE" ~
    # "DELIVERY TENSION REEL ASSEMBLY"): link as a weaker name match
    free_s = [s for s in S if not ls[s["i"]] and s["bill"] > 0 and s["disp"]]
    free_w = [w for w in W if not lw[w["i"]] and w["net"] > 0]
    best_of_s = {s["i"]: max(((nscore.get((s["i"], w["i"]), 0), w["i"]) for w in free_w), default=(0, None))
                 for s in free_s}
    best_of_w = {w["i"]: max(((nscore.get((s["i"], w["i"]), 0), s["i"]) for s in free_s), default=(0, None))
                 for w in free_w}
    for si, (sc, wi) in best_of_s.items():
        if wi is not None and sc >= 55 and best_of_w[wi][1] == si and (sidx[si]["skey"], widx[wi]["wbs"]) not in unlink \
                and near(sidx[si]["month"], widx[wi]["last"], 4):
            link(si, wi, "name", sc)

    # same value (>= ₹1 L, unique in the project) billed within a month of the forecast: link, flagged "check"
    cs = Counter(round(x["bill"]) for x in S if not ls[x["i"]] and x["disp"])
    cw = Counter(round(x["net"]) for x in W if not lw[x["i"]])
    for s in S:
        if ls[s["i"]] or not s["disp"] or s["bill"] < 1e5 or cs[round(s["bill"])] != 1:
            continue
        for w in W:
            if not lw[w["i"]] and cw[round(w["net"])] == 1 and abs(w["net"] - s["bill"]) <= _tol(s["bill"]) \
                    and near(s["month"], w["last"], 1) and (s["skey"], w["wbs"]) not in unlink:
                link(s["i"], w["i"], "value", nscore.get((s["i"], w["i"]), 0))
                break

    # connected components
    seen_s, seen_w, groups = set(), set(), []
    for x in S:
        if x["i"] in seen_s:
            continue
        GS, GW, stack = set(), set(), [("s", x["i"])]
        while stack:
            kind, y = stack.pop()
            if kind == "s" and y not in GS:
                GS.add(y); stack += [("w", z) for z in ls[y]]
            elif kind == "w" and y not in GW:
                GW.add(y); stack += [("s", z) for z in lw[y]]
        seen_s |= GS; seen_w |= GW
        groups.append((GS, GW, {k: v for k, v in how.items() if k[0] in GS}))
    for w in W:
        if w["i"] not in seen_w:
            groups.append((set(), {w["i"]}, {}))
    return groups


# ----------------------------------------------------------------------------- the book

STATUS = {
    "billed_not_disp": ("Billed in SAP, sales file says Not Disp.", "red"),
    "disp_not_billed": ("Marked Disp., no SAP bill", "red"),
    "sap_only": ("Billed in SAP, not in sales file", "amber"),
    "part_billed": ("Partly billed", "amber"),
    "value_diff": ("Billed, value differs", "amber"),
    "ok": ("Matches", "green"),
    "slipped": ("Not billed, month passed", "grey"),
    "pending": ("Not billed yet", "grey"),
    "service": ("Service bill (not a forecast item)", "grey"),
    "zero": ("₹0 row (no billing value)", "grey"),
}
ISSUES = ["billed_not_disp", "disp_not_billed", "sap_only", "part_billed", "value_diff"]
BASIS = {"manual": "linked by you", "value": "same value & month — check", "sum": "values add up", "name": "name"}


def _iso(m):
    return None if m is None or pd.isna(m) else pd.Timestamp(m).strftime("%Y-%m-%d")


def _mlabel(m):
    return "" if m is None or pd.isna(m) else pd.Timestamp(m).strftime("%b %Y")


class ReconBook:
    def __init__(self, sales_path: str, sap_path: str, today: dt.date | None = None):
        self.today = pd.Timestamp(today or dt.date.today())
        self.cur = E.month_start(self.today)
        s = E.load_sales(sales_path)
        s["bill"] = s.bill.fillna(0.0)
        s["row"] = s.index + 2                                   # Excel row number in PROJECTS
        s["pno"] = s.project.str.extract(r"(\d{4})")[0]
        s = s[s.pno.notna()].copy()
        s["key"] = [clean(n, p) for n, p in zip(s.name, s.pno)]
        up = s.name.str.upper().str.strip()
        s["skey"] = up + "|" + s.groupby(["pno", up]).cumcount().astype(str)
        self.sales = s
        self.lines, self.wbs, self.meta = load_sap(sap_path)
        self._cache: dict = {}

    def invalidate(self, pno: str | None = None):
        if pno:
            self._cache.pop(pno, None)
        else:
            self._cache.clear()

    # ------------------------------------------------------------------ one project
    def project(self, pno: str) -> dict:
        if pno in self._cache:
            return self._cache[pno]
        ov = load_overrides().get(pno, {})
        notes = ov.get("notes", {})
        srows = self.sales[self.sales.pno == pno]
        wrows = self.wbs[self.wbs.pno == pno] if len(self.wbs) else self.wbs
        out = []
        srec = dict(zip(srows.index, srows.to_dict("records")))
        wrec = dict(zip(wrows.index, wrows.to_dict("records")))
        for gi, (GS, GW, how) in enumerate(match_project(srows, wrows, ov)):
            s_recs = sorted((srec[i] for i in GS), key=lambda r: r["row"])
            w_recs = [wrec[i] for i in sorted(GW)]
            f_val = float(sum(r["bill"] for r in s_recs))
            s_net = float(sum(r["net"] for r in w_recs))
            billed = s_net > max(1.0, 0.02 * f_val)
            disp = [bool(r["dispatched"]) for r in s_recs]
            fm = [r["month"] for r in s_recs if pd.notna(r["month"])]
            f_month = max(fm) if fm else pd.NaT
            sm = [r["last"] for r in w_recs if pd.notna(r["last"])]
            s_month = E.month_start(max(sm)) if sm else pd.NaT
            if not s_recs:
                if s_net <= 0:                               # fully cancelled, no forecast row
                    continue
                st = "service" if all(r["service"] for r in w_recs) else "sap_only"
            elif f_val <= 0 and not billed:
                st = "zero"
            elif billed and not all(disp):
                st = "billed_not_disp"
            elif billed and s_net < PART_BILLED * f_val - VALUE_TOL_ABS:
                st = "part_billed"
            elif billed:
                st = "ok" if abs(s_net - f_val) <= max(VALUE_TOL * f_val, VALUE_TOL_ABS) else "value_diff"
            elif any(disp):
                st = "disp_not_billed"
            else:
                st = "slipped" if pd.notna(f_month) and f_month < self.cur else "pending"
            kinds = {v[0] for v in how.values()}
            basis = next((k for k in ("manual", "value", "sum", "name") if k in kinds), "")
            scores = [v[1] for v in how.values() if v[0] == "name"]
            out.append({
                "gid": gi, "status": st, "label": STATUS[st][0], "tone": STATUS[st][1],
                "sales": [{"skey": r["skey"], "row": int(r["row"]), "name": r["name"], "boi": r["boi"], "pm": r["pm"],
                           "bill": float(r["bill"]), "month": _iso(r["month"]), "month_label": _mlabel(r["month"]),
                           "status": "Disp." if r["dispatched"] else "Not Disp.", "dispatched": bool(r["dispatched"]),
                           "risk": r["risk"], "remarks": r["remarks"], "note": notes.get(r["skey"], "")}
                          for r in s_recs],
                "sap": [{"wbs": r["wbs"], "desc": r["desc"], "names": [n for n in r["names"][1:] if n != r["desc"]][:4],
                         "net": r["net"], "gross": r["gross"], "cancelled": r["cancelled"], "cur": r["cur"],
                         "net_tc": r["net_tc"], "first": _iso(r["first"]), "last": _iso(r["last"]),
                         "last_label": _mlabel(r["last"]), "n": int(r["n"]), "lines": r["lines"],
                         "note": notes.get(f"SAP:{r['wbs']}", "")}
                        for r in w_recs],
                "f_value": f_val, "s_value": s_net, "diff": s_net - f_val,
                "f_month": _iso(f_month), "f_month_label": _mlabel(f_month),
                "s_month": _iso(s_month), "s_month_label": _mlabel(s_month),
                "month_diff": _months_between(f_month, s_month) if billed and s_recs else None,
                "score": round(min(scores)) if scores else None, "basis": basis,
                "basis_label": BASIS.get(basis, ""),
            })
        order = {k: i for i, k in enumerate(STATUS)}
        out.sort(key=lambda g: (order[g["status"]], g["sales"][0]["row"] if g["sales"] else 10 ** 6))
        res = {"pno": pno, "groups": out, "summary": self._summ(pno, out, srows, wrows)}
        self._cache[pno] = res
        return res

    def _summ(self, pno, groups, srows, wrows) -> dict:
        cnt = {k: 0 for k in STATUS}
        val = {k: 0.0 for k in STATUS}
        for g in groups:
            cnt[g["status"]] += 1
            if g["status"] in ("billed_not_disp", "sap_only", "service"):
                val[g["status"]] += g["s_value"]
            elif g["status"] in ("part_billed", "value_diff"):
                val[g["status"]] += abs(g["diff"])
            else:
                val[g["status"]] += g["f_value"]
        sap_info = self.meta["projects"].get(pno, {})
        name = srows.project.iloc[0] if len(srows) else f"{pno} - {str(sap_info.get('sap_desc', '')).strip()}"
        return {
            "pno": pno, "project": name, "in_sales": bool(len(srows)),
            "pm": ", ".join(sorted({x for x in srows.pm if x})),
            "customer": re.sub(r"^\s*M\s*/\s*S\.?\s*", "", str(sap_info.get("customer", "") or "")).split("/")[0].strip(),
            "forecast": float(srows.bill.sum()), "marked_disp": float(srows.loc[srows.dispatched, "bill"].sum()),
            "sap_billed": float(wrows.net.sum()) if len(wrows) else 0.0,
            "rows": int(len(srows)), "wbs": int(len(wrows)),
            "counts": cnt, "values": val,
            "issues": int(sum(cnt[k] for k in ISSUES)),
            "red": int(cnt["billed_not_disp"] + cnt["disp_not_billed"]),
            "matched": int(sum(1 for g in groups if g["sales"] and g["sap"])),
        }

    # ------------------------------------------------------------------ all projects
    def pnos(self) -> list[str]:
        return sorted(set(self.sales.pno) | set(self.wbs.pno))

    def overview(self) -> dict:
        rows = [self.project(p)["summary"] for p in self.pnos()]
        return {
            "projects": rows,
            "counts": {k: sum(r["counts"][k] for r in rows) for k in STATUS},
            "values": {k: sum(r["values"][k] for r in rows) for k in STATUS},
            "labels": {k: v[0] for k, v in STATUS.items()}, "tones": {k: v[1] for k, v in STATUS.items()},
            "issues": ISSUES,
            "sap_from": _iso(self.meta["from"]), "sap_to": _iso(self.meta["to"]),
            "sap_unlinked_rows": self.meta["unlinked_rows"], "sap_unlinked_inr": self.meta["unlinked_inr"],
            "today": _iso(self.today),
        }

    def issues(self) -> list[dict]:
        """Every mismatch group across projects (light rows for the cross-project list)."""
        out = []
        for p in self.pnos():
            pr = self.project(p)
            for g in pr["groups"]:
                if g["status"] not in ISSUES:
                    continue
                out.append({
                    "pno": p, "project": pr["summary"]["project"], "gid": g["gid"], "status": g["status"],
                    "label": g["label"], "tone": g["tone"],
                    "sales": " + ".join(s["name"] for s in g["sales"]), "rows": [s["row"] for s in g["sales"]],
                    "f_value": g["f_value"] if g["sales"] else None, "f_month_label": g["f_month_label"],
                    "sap": " + ".join(w["desc"] for w in g["sap"]), "wbs": [w["wbs"] for w in g["sap"]],
                    "s_value": g["s_value"] if g["sap"] else None, "s_month_label": g["s_month_label"],
                    "diff": g["diff"], "pm": ", ".join(sorted({s["pm"] for s in g["sales"] if s["pm"]})),
                    "basis_label": g["basis_label"],
                })
        return out

    def side(self, pno: str) -> dict:
        """One project as three lists: in both / only in SAP / only in the sales file."""
        pr = self.project(pno)
        both, sap_only, sales_only = [], [], []
        for g in pr["groups"]:
            base = {"gid": g["gid"], "status": g["status"], "tone": g["tone"], "label": g["label"]}
            if g["sales"] and g["sap"]:
                last = max((w["last"] for w in g["sap"] if w["last"]), default=None)
                both.append({**base,
                             "sales": [{"name": s["name"], "row": s["row"], "status": s["status"]} for s in g["sales"]],
                             "sap": [{"desc": w["desc"], "wbs": w["wbs"]} for w in g["sap"]],
                             "f_value": g["f_value"], "f_month": g["f_month_label"],
                             "s_value": g["s_value"], "s_date": last, "bills": sum(w["n"] for w in g["sap"]),
                             "diff": g["diff"], "month_diff": g["month_diff"]})
            elif g["sap"]:
                for w in g["sap"]:
                    sap_only.append({**base, "desc": w["desc"], "wbs": w["wbs"], "value": w["net"],
                                     "date": w["last"], "first": w["first"], "bills": w["n"]})
            else:
                for s in g["sales"]:
                    sales_only.append({**base, "name": s["name"], "row": s["row"], "value": s["bill"],
                                       "month": s["month"], "month_label": s["month_label"],
                                       "disp": s["status"], "boi": s["boi"]})
        both.sort(key=lambda x: x["sales"][0]["row"])
        sap_only.sort(key=lambda x: (x["date"] or "9999", x["desc"]))
        sales_only.sort(key=lambda x: x["row"])
        return {"summary": pr["summary"], "both": both, "sap_only": sap_only, "sales_only": sales_only}

    def matrix(self, pno: str) -> dict:
        """One row per assembly. In both → SAP billed value & billing date; otherwise its own value & month."""
        d = self.side(pno)
        rows = []
        for x in d["both"]:
            rows.append({"gid": x["gid"], "kind": "both", "status": x["status"], "tone": x["tone"], "label": x["label"],
                         "name": " + ".join(s["name"] for s in x["sales"]),
                         "other": " + ".join(w["desc"] for w in x["sap"]),
                         "ref": ", ".join(str(s["row"]) for s in x["sales"]) + " | " + ", ".join(w["wbs"] for w in x["sap"]),
                         "value": x["s_value"], "date": x["s_date"], "f_value": x["f_value"], "f_month": x["f_month"],
                         "disp": ", ".join(sorted({s["status"] for s in x["sales"]})), "bills": x["bills"]})
        for x in d["sap_only"]:
            rows.append({"gid": x["gid"], "kind": "sap", "status": x["status"], "tone": x["tone"], "label": x["label"],
                         "name": x["desc"], "other": "", "ref": x["wbs"], "value": x["value"], "date": x["date"],
                         "f_value": None, "f_month": "", "disp": "", "bills": x["bills"]})
        for x in d["sales_only"]:
            rows.append({"gid": x["gid"], "kind": "sales", "status": x["status"], "tone": x["tone"], "label": x["label"],
                         "name": x["name"], "other": "", "ref": str(x["row"]), "value": x["value"], "date": x["month"],
                         "f_value": x["value"], "f_month": x["month_label"], "disp": x["disp"], "bills": 0})
        for r in rows:
            r["month"] = r["date"][:7] if r["date"] else ""
        rows.sort(key=lambda r: (r["month"] or "9999", {"both": 0, "sap": 1, "sales": 2}[r["kind"]], r["name"]))
        months = sorted({r["month"] for r in rows if r["month"]})
        if months:   # fill gaps so the columns are continuous
            rng = pd.period_range(months[0], months[-1], freq="M")
            months = [str(m) for m in rng]
        return {"summary": d["summary"], "rows": rows, "months": months}

    def fy_kpis(self) -> dict:
        """Indian FY (Apr-Mar) headline numbers.
        forecast: sales-file rows from the current month to the FY end; actual: SAP billing from 1 Apr to today;
        expected: SAP billed + forecast rows of the window that are not yet Disp. (Disp. ones are already in SAP)."""
        t = self.today
        fy0 = pd.Timestamp(t.year if t.month >= 4 else t.year - 1, 4, 1)
        fy1 = pd.Timestamp(fy0.year + 1, 3, 1)                     # month start of March
        s = self.sales
        win = s[(s.month >= self.cur) & (s.month <= fy1)]
        todo = win[~win.dispatched]
        slipped = s[(s.month < self.cur) & (s.month >= fy0) & ~s.dispatched & (s.bill > 0)]
        L = self.lines[(self.lines.date >= fy0) & (self.lines.date <= t + pd.Timedelta(days=1))]
        A = self.meta["all_lines"]
        A = A[(A.date >= fy0) & (A.date <= t + pd.Timedelta(days=1))]
        per_wbs = L.groupby(["pno", "wbs"]).inr.sum()
        per_pno = L.groupby("pno").inr.sum()
        actual = float(A.inr.sum())
        todo_v = float(todo.bill.sum())
        fy = f"{fy0.year}-{str(fy0.year + 1)[2:]}"
        return {
            "fy": fy, "fy_start": _iso(fy0), "today": _iso(t), "win_from": _iso(self.cur), "win_to": _iso(fy1),
            "forecast": {"value": float(win.bill.sum()), "disp_value": float(win.loc[win.dispatched, "bill"].sum()),
                         "projects": int(win.loc[win.bill > 0, "pno"].nunique()),
                         "to_dispatch": int((todo.bill > 0).sum()), "to_dispatch_value": todo_v,
                         "to_dispatch_zero": int((todo.bill <= 0).sum()),
                         "slipped": int(len(slipped)), "slipped_value": float(slipped.bill.sum())},
            "actual": {"value": actual, "project_value": float(L.inr.sum()),
                       "projects": int((per_pno > 0).sum()), "assemblies": int((per_wbs > 0).sum()),
                       "lines": int(len(A))},
            "expected": {"value": actual + todo_v, "sap": actual, "forecast": todo_v},
        }

    def side_all(self) -> list[dict]:
        out = []
        for p in self.pnos():
            d = self.side(p)
            if d["both"] or d["sap_only"] or d["sales_only"]:
                out.append(d)
        return out

    def candidates(self, pno: str, skey: str | None = None, wbs: str | None = None) -> list[dict]:
        """For a manual link: the other side's items, best name match first."""
        srows = self.sales[self.sales.pno == pno]
        wrows = self.wbs[self.wbs.pno == pno]
        if skey:
            s = srows[srows.skey == skey].iloc[0]
            sc = _score_matrix([s.key], wrows["keys"].tolist())[0]
            c = [{"wbs": w["wbs"], "desc": w["desc"], "net": w["net"], "last": _mlabel(w["last"]), "score": round(float(x)),
                  "fit": _value_fit(float(s.bill), w["net"])}
                 for w, x in zip(wrows.to_dict("records"), sc)]
        else:
            w = wrows[wrows.wbs == wbs].iloc[0]
            sc = _score_matrix(srows.key.tolist(), [w["keys"]])[:, 0]
            c = [{"skey": s["skey"], "name": s["name"], "bill": float(s["bill"]), "month": _mlabel(s["month"]),
                  "status": "Disp." if s["dispatched"] else "Not Disp.", "score": round(float(x)),
                  "fit": _value_fit(float(s["bill"]), float(w.net))}
                 for s, x in zip(srows.to_dict("records"), sc)]
        # name first, a near-equal value lifts an item up the list
        return sorted(c, key=lambda x: -(x["score"] + (25 if x["fit"] >= 0.98 else 10 if x["fit"] >= 0.9 else 0)))

    def table(self, pnos: list[str] | None = None, statuses: list[str] | None = None) -> pd.DataFrame:
        recs = []
        for p in pnos or self.pnos():
            pr = self.project(p)
            for g in pr["groups"]:
                if statuses and g["status"] not in statuses:
                    continue
                recs.append({
                    "Project": pr["summary"]["project"], "Status": g["label"],
                    "Sales row(s)": ", ".join(str(s["row"]) for s in g["sales"]),
                    "Sales assembly": " + ".join(s["name"] for s in g["sales"]),
                    "Sales status": ", ".join(sorted({s["status"] for s in g["sales"]})),
                    "Forecast value": g["f_value"] if g["sales"] else None, "Forecast month": g["f_month_label"],
                    "SAP WBS": ", ".join(w["wbs"] for w in g["sap"]),
                    "SAP item": " + ".join(w["desc"] for w in g["sap"]),
                    "SAP net billed": g["s_value"] if g["sap"] else None, "SAP bill month": g["s_month_label"],
                    "Difference (SAP - forecast)": g["diff"] if g["sap"] and g["sales"] else None,
                    "Months late (+) / early (-)": g["month_diff"],
                    "Matched by": (f"name ({g['score']})" if g["basis"] == "name" and g["score"] is not None
                                   else g["basis_label"]),
                    "PM": ", ".join(sorted({s["pm"] for s in g["sales"] if s["pm"]})),
                    "Note": "; ".join(x["note"] for x in g["sales"] + g["sap"] if x.get("note")),
                })
        return pd.DataFrame(recs)
