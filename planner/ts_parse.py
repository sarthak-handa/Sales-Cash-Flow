import re, sys, collections
CODE = r"(?:[SBO]|-)"
UNITS = r"(?:LOT|LOTS|SET|SETS|NOS?|NO\.|PCS|PC|MTR|M|KG|UNIT|UNITS)"
row_re = re.compile(rf"^(?P<sn>\d+(?:\.\d+)?)\s+(?P<desc>.+?)\s+(?:(?P<qty>\d+(?:\.\d+)?)\s+(?P<unit>{UNITS})\s+)?(?P<codes>(?:{CODE}\s+){{0,7}}{CODE})\s*$", re.I)
hdr_re = re.compile(r"^(?P<sn>\d+(?:\.\d+)?)\s+(?P<desc>[A-Z][^a-z]{3,})$")
NOISE = re.compile(r"THIS DOCUMENT|AUTHORI[SZ]ATION|OFFER NO|Page \d+ of|=====PAGE|^BSD|^Commissioning, SCO|^S ?NO\b|^S N\b|DIVISION LIST|YD/[A-Z]+/", re.I)
def parse(path):
    lines = open(path, encoding="utf-8").read().splitlines()
    s = next((i for i, l in enumerate(lines) if "DIVISION LIST" in l.upper()), None)
    if s is None: return []
    out, pending, sec = [], "", ""
    for l in lines[s:]:
        if re.match(r"^CHAPTER\s*-?\s*\d", l.strip(), re.I) and out: break
        t = l.strip()
        if re.search(r"DESCRIPTION\s+QTY\s+UNIT", t, re.I): sec = ""; pending = ""; continue
        if not t or NOISE.search(t): continue
        if re.match(r"^[A-H]\s+[A-Z]", t) and not re.match(r"^[A-H]\s+\d", t):   # lettered general sections
            sec = t; pending = ""; continue
        m = row_re.match(t)
        if m and len(m["codes"].split()) >= 1 and not m["desc"].isdigit():
            codes = m["codes"].split()
            desc = (pending + " " + m["desc"]).strip() if pending else m["desc"]
            out.append(dict(sr=m["sn"], name=desc, qty=m["qty"] or "", unit=m["unit"] or "", codes=codes,
                            sup=(codes[3] if len(codes) >= 8 else ("S" if codes and codes[0]=="S" and len(codes)>=4 else "?")),
                            sec=sec, group=False)); pending = ""; continue
        m = hdr_re.match(t)
        if m:
            out.append(dict(sr=m["sn"], name=m["desc"].strip(), qty="", unit="", codes=[], sup="", sec=sec, group=True)); pending=""; continue
        pending = t if len(t) < 90 else ""
    return out
if __name__ == "__main__":
    for p in sys.argv[1:]:
        rows = parse(p)
        eq = [r for r in rows if not r["sec"] or r["sec"][:1] not in "ABCDEFGH"]
        print("####", p.split("txt_")[-1], "rows", len(rows), "equipment rows", len(eq),
              "sup:", collections.Counter(r["sup"] for r in eq if not r["group"]))
        for r in eq[:12]: print("  ", r["sr"], r["name"][:60], r["qty"], r["unit"], r["sup"], "GROUP" if r["group"] else "")
