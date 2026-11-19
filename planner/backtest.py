"""Backtest: re-plan 2611 and 2618 without letting the engine see them, compare with the planner's plan."""
import datetime as dt
import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
import engine as E  # noqa: E402
E.LEAD_FLOOR = "nofloor" not in sys.argv
E.USE_BEST_REF = "nobest" not in sys.argv
E.USE_LOTS = "nolots" not in sys.argv
for _a in sys.argv:
    if _a.startswith("ref="):
        E.REF_MODE = _a[4:]
E.REF_GROUP_SUM = "nogroup" not in sys.argv

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import datasource  # noqa: E402
SALES = datasource.sales_path()[0]   # same workbook the app reads (config.json)

# (pno, project, family, PO value, PO date, delivery, window override)
# 2621 / 2630 (CRM): PO not available -> PO value = the project total in the sales file (incl. electrical,
# HT motors, spares, which are not in the LoE), window = his first..last month,
# so only "which rows / how much / in what order" is tested, not the window itself.
CASES = [
    ("2611", "2611-APL-CGL", "CGL", 59.71e7, dt.date(2026, 3, 7), dt.date(2026, 11, 10), None),
    ("2618", "2618-APL-CGL", "CGL", 79.80e7, dt.date(2026, 3, 13), dt.date(2027, 3, 15), None),
    ("2621", "2621-PROMPT-6HI CRM", "CRM", 36.50e7, dt.date(2026, 7, 1), dt.date(2027, 4, 1), (dt.date(2027, 4, 1), dt.date(2027, 9, 1))),
    ("2630", "2630-JSW-6HI CRM", "CRM", 32.00e7, dt.date(2026, 4, 1), dt.date(2027, 1, 1), (dt.date(2026, 11, 1), dt.date(2027, 5, 1))),
]
only = [a for a in sys.argv[1:] if a.isdigit()]

for pno, proj, fam, value, po_date, deliv, win in CASES:
    if only and pno not in only:
        continue
    kb = E.load_knowledge(SALES, ROOT, exclude={pno})
    E.exclude_refs = {pno}
    loe = E.parse_loe(os.path.join(ROOT, f"{pno} planning.xlsx"))
    truth = E.load_reference_plans(ROOT)
    truth = truth[truth.project == pno].set_index("row")
    po = E.POInfo(project=proj, family=fam, po_value=value, po_date=po_date, delivery_date=deliv)
    if win:
        po.start_month, po.end_month = win
    # backtest as of PO date (do not shift window to today)
    E_today = pd.Timestamp.today
    pd.Timestamp.today = staticmethod(lambda: pd.Timestamp(po_date))
    try:
        df = E.plan(loe, po, kb)
    finally:
        pd.Timestamp.today = E_today
    df = df.join(truth[["bill", "month", "billed"]], on="row")
    lots = df[df.section == "NOT IN LoE"]
    if len(lots):
        print("  lots added:", ", ".join(f"{n} {v/1e7:.2f} Cr" for n, v in zip(lots.name, lots.billing_value)))
    df = df[df.section != "NOT IN LoE"]
    t, p = df.billed.fillna(False).astype(bool), df.billable
    tp = (t & p).sum()
    print(f"\n===== {pno}: rows {len(df)}, planner billed {t.sum()}, engine billed {p.sum()}")
    print(f"  billable-row agreement: precision {tp/p.sum():.0%}, recall {tp/t.sum():.0%}")
    vw = df.loc[t, "bill"].sum()
    print(f"  share of planner's value on rows the engine also bills: {df.loc[t & p, 'bill'].sum()/vw:.0%}")
    both = df[t & p]
    r = both.billing_value / both.bill
    print(f"  value ratio engine/planner on common rows: median {r.median():.2f}, within ±30% {((r>.7)&(r<1.3)).mean():.0%}, within ±50% {((r>.5)&(r<1.5)).mean():.0%}")
    md = ((both.dispatch_month.dt.year*12+both.dispatch_month.dt.month) - (both.month.dt.year*12+both.month.dt.month))
    print(f"  dispatch month diff on common rows: exact {(md==0).mean():.0%}, within ±1 month {(md.abs()<=1).mean():.0%}")
    a = df[p].groupby("dispatch_month").billing_value.sum() / 1e7
    b = df[t].groupby("month").bill.sum() / 1e7
    m = pd.concat([a.rename("engine Cr"), b.rename("planner Cr")], axis=1).fillna(0).round(2)
    m.index = m.index.strftime("%b %Y")
    print(m.to_string())
    miss = df[t & ~p].sort_values("bill", ascending=False)[["sr", "name", "bill", "reason"]].head(8)
    extra = df[~t & p].sort_values("billing_value", ascending=False)[["sr", "name", "billing_value", "reason"]].head(8)
    print("  biggest planner rows the engine skipped:\n" + miss.to_string())
    print("  rows engine billed that planner didn't:\n" + extra.to_string())
