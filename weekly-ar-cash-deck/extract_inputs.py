#!/usr/bin/env python3
"""
extract_inputs.py — builds the weekly input workbook (weekly_inputs.xlsx) from the four source files:

  1. AR dashboard           (AR_PRE.xlsx            : sheets 'Executive Summary' + 'Comparatives')
  2. Cash flow workbook     (CASH_FLOW_ANALYSIS...  : 'EXEC Summary', 'Approved vs Tracking log', 'Priority list', 'LC Ongoing')
  3. LC situation           (LCs_SITUATION...       : sheet 'LC avec echéance')
  4. Supplier invoices      (Supplier_Invoices...   : sheet '4411 - Parked - Downpayment')

Usage:
  python3 extract_inputs.py --ar AR_PRE.xlsx --cash CASH_FLOW.xlsx --lc LCs.xlsx --ap Supplier_Invoices.xlsx \
          --as-of 2026-10-21 --ar-date 2026-10-14 --out inputs/weekly_inputs.xlsx

All amounts are written in MMAD (millions of MAD). Every tab of the output workbook can be edited by hand
before running build_deck.py (e.g. to adjust the key messages).
"""
import argparse, datetime as dt, json, re, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------- helpers
BU_CANON = {
    "advisory": "Advisory", "ads": "Advisory",
    "am": "AM", "amt": "AM",
    "building & capabilities": "Building & Capabilities", "b&c": "Building & Capabilities",
    "building & infrastructure": "Building & Infrastructure", "bin": "Building & Infrastructure", "b&i": "Building & Infrastructure",
    "energy operations": "Energy Operations", "ene": "Energy Operations", "energy": "Energy Operations",
    "fertilizers & chemicals": "Fertilizers & Chemicals", "fec": "Fertilizers & Chemicals",
    "global mining": "Global Mining", "gmm": "Global Mining",
    "jesa technologies": "JESA Technologies", "jet": "JESA Technologies",
    "sp2m ad hoc program": "SP2M Ad Hoc Program", "sp2m": "SP2M Ad Hoc Program",
    "water & environnement": "Water & Environment", "water & environement": "Water & Environment",
    "water & environment": "Water & Environment", "wen": "Water & Environment",
    "power to x": "Power to X",
}
BU_ORDER = ["Water & Environment", "Fertilizers & Chemicals", "SP2M Ad Hoc Program", "AM", "Energy Operations",
            "Global Mining", "Building & Infrastructure", "JESA Technologies", "Advisory", "Building & Capabilities", "Power to X"]

def canon_bu(name):
    if name is None: return None
    k = str(name).strip().lower()
    k = re.sub(r"\s+", " ", k)
    if k in BU_CANON: return BU_CANON[k]
    for key, v in BU_CANON.items():
        if k.startswith(key): return v
    return None

def num(v, scale=1.0):
    """float or 0; scale divides (1e6 to convert MAD -> MMAD)."""
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v) / scale
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "n/a", "NA", "na", "XXXX"): return 0.0
    try: return float(s) / scale
    except ValueError: return 0.0

def find_row(rows, text, col=None, start=0, exact=False):
    """index of first row whose cell (col or any) matches text (case-insensitive startswith)."""
    t = text.lower()
    for i in range(start, len(rows)):
        cells = [rows[i][col]] if col is not None else rows[i]
        for c in cells:
            if c is None: continue
            s = str(c).strip().lower()
            if (s == t) if exact else s.startswith(t):
                return i
    raise KeyError(f"label not found: {text!r}")

def load(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    return [list(r) for r in ws.iter_rows(values_only=True)]

# ----------------------------------------------------------------------------- AR dashboard
def extract_ar(path):
    out = {}
    cmp_ = load(path, "Comparatives")
    exe = load(path, "Executive Summary")

    # baseline dates
    r = cmp_[find_row(cmp_, "Previous baseline saved")]
    def xl_date(v):
        if isinstance(v, (dt.date, dt.datetime)): return v.date() if isinstance(v, dt.datetime) else v
        if isinstance(v, (int, float)): return (dt.date(1899, 12, 30) + dt.timedelta(days=int(v)))
        return None
    out["prev_baseline_date"] = xl_date(r[2]); out["month_start_date"] = xl_date(r[5])

    # KEY FIGURES (B=metric C=MS D=PW E=Cur)
    i = find_row(cmp_, "Metric", col=1) + 1
    kpi = []
    while cmp_[i][1]:
        kpi.append([str(cmp_[i][1]).strip(), num(cmp_[i][2]), num(cmp_[i][3]), num(cmp_[i][4])]); i += 1
    out["kpi"] = kpi

    # 1. AR BY BU (blocks of 5 cols: MS, PW, Cur, dWoW, dMTD) for Advance/Unbilled/Unpaid/Total
    i = find_row(cmp_, "1. AR BY BUSINESS UNIT") + 3
    ar_bu = []
    while cmp_[i][1] and str(cmp_[i][1]).strip().lower() != "total":
        bu = canon_bu(cmp_[i][1]); row = cmp_[i]
        for j, comp in enumerate(["Advance", "Unbilled", "Unpaid", "Total AR"]):
            b = 2 + 5 * j
            ar_bu.append([bu, comp, num(row[b]), num(row[b + 1]), num(row[b + 2])])
        i += 1
    out["ar_bu"] = ar_bu

    # 2. UNBILLED BY CATEGORY (At risk / JESA / OCP / SS postponed)
    i = find_row(cmp_, "2. UNBILLED BY CATEGORY") + 3
    owner = []
    while cmp_[i][1] and str(cmp_[i][1]).strip().lower() != "total":
        bu = canon_bu(cmp_[i][1]); row = cmp_[i]
        for j, comp in enumerate(["At risk", "Under JESA", "Under OCP", "SS postponed"]):
            b = 2 + 5 * j
            owner.append([bu, comp, num(row[b]), num(row[b + 1]), num(row[b + 2])])
        i += 1
    out["unbilled_owner"] = owner

    # 5. AGING BY YEAR
    i = find_row(cmp_, "5. AR AGING BY YEAR") + 2
    aging = []
    while cmp_[i][1] and str(cmp_[i][1]).strip().lower() != "total":
        aging.append([str(cmp_[i][1]).strip(), num(cmp_[i][2]), num(cmp_[i][3]), num(cmp_[i][4])]); i += 1
    out["aging"] = aging

    # 6. PAYABLES & LC BY BU (PW, Cur, d) x 4 buckets
    i = find_row(cmp_, "6. PAYABLES & ATTACHMENTS BY BU") + 3
    ap = []
    while cmp_[i][1] and str(cmp_[i][1]).strip().lower() != "total":
        bu = canon_bu(cmp_[i][1]); row = cmp_[i]
        ap.append([bu] + [num(row[2 + 3 * j + k]) for j in range(4) for k in (0, 1)])
        i += 1
    out["ap_bu"] = ap

    # 7. UNBILLED BY STATUS 2
    i = find_row(cmp_, "7. UNBILLED BY STATUS 2") + 4
    st = []; group = None
    while True:
        row = cmp_[i]; lab = row[1]
        if lab is None: i += 1; continue
        s = str(lab).strip()
        if s.lower().startswith("total"): break
        if s in ("OCP", "JESA", "At Risk"): group = s
        elif "subtotal" in s.lower(): pass
        elif s.lower().startswith("check"): pass
        elif s.lower().startswith("other / unlisted"): st.append(["Other", s, num(row[2]), num(row[3]), num(row[4])])
        else: st.append([group, s, num(row[2]), num(row[3]), num(row[4])])
        i += 1
    out["unbilled_status"] = st

    # 8. UNPAID BY STATUS
    i = find_row(cmp_, "8. UNPAID BY STATUS") + 3
    up = []
    while cmp_[i][1] and str(cmp_[i][1]).strip().lower() != "total":
        up.append([str(cmp_[i][1]).strip(), num(cmp_[i][2]), num(cmp_[i][3]), num(cmp_[i][4])]); i += 1
    out["unpaid_status"] = up

    # 9. PAYMENTS RECEIVED BY BU
    i = find_row(cmp_, "9. PAYMENTS RECEIVED BY BU") + 3
    col = []
    while cmp_[i][1] and str(cmp_[i][1]).strip().lower() != "total":
        row = cmp_[i]
        col.append([canon_bu(row[1]), num(row[2]), num(row[3]), num(row[6]), num(row[7]), num(row[8]), int(num(row[12]))]); i += 1
    out["collections_bu"] = col

    # Executive summary: labor / non-labor key figures (row after 'NEW AR' header)
    i = find_row(exe, "NEW AR", col=1)
    r1, r2 = exe[i + 2], exe[i + 3]
    out["labor_split"] = {
        "total_labor": num(r1[5], 1e6), "total_nonlabor": num(r2[5], 1e6),
        "unbilled_labor": num(r1[11], 1e6), "unbilled_nonlabor": num(r2[11], 1e6),
        "unpaid_labor": num(r1[14], 1e6), "unpaid_nonlabor": num(r2[14], 1e6),
    }
    # OCP vs NON OCP
    i = find_row(exe, "OCP vs NON OCP")
    out["ocp_split"] = {str(exe[k][1]).strip(): num(exe[k][5], 1e6) for k in (i + 2, i + 3)}
    # AR by BU incl. unrecognized (current)
    i = find_row(exe, "1. AR by Business Unit") + 2
    unrec = {}
    while exe[i][1] and str(exe[i][1]).strip().lower() != "total ar":
        unrec[canon_bu(exe[i][1])] = num(exe[i][6], 1e6); i += 1
    out["unrecognized_bu"] = unrec
    # 7. Unrecognized suppliers table
    i = find_row(exe, "7. Unrecognized") + 3
    ut = []
    while exe[i][1] and str(exe[i][1]).strip().lower() != "total":
        row = exe[i]; bu = canon_bu(row[1])
        if bu: ut.append([bu] + [num(row[c], 1e6) for c in range(2, 15)])
        i += 1
    out["unrecognized_table"] = ut
    return out

# ----------------------------------------------------------------------------- Cash flow workbook
def extract_cash(path):
    out = {}
    ex = load(path, "EXEC Summary")
    def series(label, start=0, col=2):
        i = find_row(ex, label, col=col, start=start)
        r = ex[i]
        return [num(r[c], 1e6) for c in range(3, 9)], i   # Jul, Aug, Sep, Oct actual, Oct ongoing, Cumulative
    r = ex[find_row(ex, "Data as of", col=2)]
    out["data_as_of"] = r[4] if isinstance(r[4], dt.datetime) else None
    basis = ex[find_row(ex, "Basis in force", col=2)]
    out["basis"] = str([c for c in basis if c and str(c).strip() != "Basis in force"][0])
    s1 = find_row(ex, "1.   CASH POSITION BY MONTH", col=2)
    lines = [("Procurement collections (OCP payment report)", "Procurement collections"),
             ("Dividends / PROC compensable", "+ Cash IN from dividends"),
             ("Total cash IN (procurement)", "Total Cash IN (procurement only)"),
             ("Cash OUT suppliers (by tier)", "Cash OUT — suppliers by tier"),
             ("Cash OUT LC payments", "Cash OUT — LC payments"),
             ("Cash OUT customs", "Cash OUT — customs"),
             ("Exclusions added back (overheads)", "Cash OUT — exclusions added back"),
             ("Total cash OUT (procurement)", "Total Cash OUT (procurement only)"),
             ("NET position (cash IN + cash OUT)", "NET POSITION")]
    monthly = []
    for lab, anchor in lines:
        vals, _ = series(anchor, start=s1)
        monthly.append([lab] + vals)
    s3 = find_row(ex, "3.   PSR (LABOR)", col=2)
    for lab, anchor in [("PSR collections (labor)", "Total PSR (Labor) collections"),
                        ("PSR provisions (AM ring-fence + W&E reserve)", "= Total provisions"),
                        ("NET PSR after provisions", "= PSR collections available")]:
        vals, _ = series(anchor, start=s3); monthly.append([lab] + vals)
    out["cash_monthly"] = monthly

    # Section 5 by BU
    s5 = find_row(ex, "5.   PROCUREMENT NET BALANCE BY BU", col=2) + 2
    bu = []
    i = s5
    while ex[i][2] and not str(ex[i][2]).strip().lower().startswith("total"):
        row = ex[i]; b = canon_bu(row[2])
        if b: bu.append([b] + [num(row[c], 1e6) for c in range(3, 12)])
        i += 1
    out["cash_bu"] = bu
    # Section 4B Oct supplier payments by BU (actual to date) and 4A Oct collections
    def block(label):
        i = find_row(ex, label, col=2) + 1; d = {}
        while ex[i][2] and not str(ex[i][2]).strip().lower().startswith("total"):
            b = canon_bu(ex[i][2])
            if b: d[b] = [num(ex[i][c], 1e6) for c in range(3, 9)]
            i += 1
        return d
    out["bu_cash_in"] = block("A.  Cash IN")
    out["bu_cash_out"] = block("B.  Cash OUT")
    out["bu_lc"] = block("C.  Cash OUT")
    out["bu_customs"] = block("D.  Cash OUT")

    # Approved vs tracking log
    try:
        tl = load(path, [s for s in openpyxl.load_workbook(path, read_only=True).sheetnames if s.startswith("Approved vs Tracking log")][0])
        i = find_row(tl, "Status vs tracking log") + 1; rows = []
        while tl[i][0] and not str(tl[i][0]).lower().startswith("total"):
            rows.append([str(tl[i][0]).strip(), int(num(tl[i][1])), abs(num(tl[i][2], 1e6))]); i += 1
        rows.append(["Total approved supplier payments", int(num(tl[i][1])), abs(num(tl[i][2], 1e6))])
        out["approved_status"] = rows
    except Exception as e:
        out["approved_status"] = []; print("WARN tracking log:", e)
    # Priority list
    try:
        pl = load(path, [s for s in openpyxl.load_workbook(path, read_only=True).sheetnames if s.startswith("Priority list")][0])
        i = find_row(pl, "Total — approved priority list")
        out["priority_total"] = num(pl[i][5], 1e6); out["priority_carried"] = num(pl[i + 2][5], 1e6)
        i = find_row(pl, "Priority", col=0, exact=True) + 1; pri = collections.defaultdict(float)
        while pl[i][0]:
            k = str(pl[i][0]).split("—")[0].strip()
            if k.startswith("P"): pri[k] += num(pl[i][5], 1e6)
            i += 1
        out["priority_by_class"] = dict(pri)
    except Exception as e:
        out["priority_total"] = 0; out["priority_carried"] = 0; out["priority_by_class"] = {}; print("WARN priority:", e)
    # LC Oct maturities counted as approved
    try:
        lo = load(path, [s for s in openpyxl.load_workbook(path, read_only=True).sheetnames if s.startswith("LC Ongoing")][0])
        i = find_row(lo, "Business Unit", col=15, exact=True) + 1; d = {}
        while lo[i][15] and str(lo[i][15]).strip().lower() != "total":
            d[canon_bu(lo[i][15])] = num(lo[i][16], 1e6); i += 1
        out["lc_oct_maturity_bu"] = d; out["lc_oct_maturity_total"] = num(lo[i][16], 1e6)
    except Exception as e:
        out["lc_oct_maturity_bu"] = {}; out["lc_oct_maturity_total"] = 0; print("WARN LC ongoing:", e)
    return out

# ----------------------------------------------------------------------------- LC situation
PROJ_TO_BU = {"FEC": "Fertilizers & Chemicals", "WEN": "Water & Environment", "ENE": "Energy Operations", "AMT": "AM"}
def extract_lc(path):
    rows = load(path, "LC avec echéance")
    hdr = [str(h).replace("\n", " ").strip() if h else "" for h in rows[0]]
    month_cols = [(i, h) for i, h in enumerate(hdr) if re.match(r"^20\d\d-\d\d$", h)]
    lines = []
    for r in rows[1:]:
        if not r[0] or str(r[0]).strip() == "": continue
        proj = str(r[3]).strip() if r[3] else ""
        bu = PROJ_TO_BU.get(proj.split("-")[0].strip(), "SP2M Ad Hoc Program" if proj.startswith("SP2M") else "Other")
        ref = str(r[4]).strip() if r[4] else ""
        opened = "Yes" if isinstance(r[9], (dt.date, dt.datetime)) else "No"   # LC opening date present
        exp = r[5]
        exp_s = exp.strftime("%Y-%m-%d") if isinstance(exp, (dt.date, dt.datetime)) else (str(exp) if exp else "")
        enc = num(r[12], 1e6)
        lines.append({
            "project": proj, "bu": bu, "supplier": str(r[6]).strip() if r[6] else "", "ref": ref, "expiry": exp_s,
            "ccy": str(r[8]).strip() if r[8] else "", "face_fx": num(r[7]), "encours_fx": num(r[11]), "encours_mad": enc,
            "action": str(r[14]).strip() if r[14] else "", "comment": str(r[13]).strip()[:120] if r[13] else "",
            "opened": opened, "part_2026": num(r[15], 1e6), "part_2027": num(r[16], 1e6),
            "months": {h: num(r[i], 1e6) for i, h in month_cols},
        })
    return {"lines": lines, "months": [h for _, h in month_cols]}

# ----------------------------------------------------------------------------- Supplier invoices (AP)
def extract_ap(path, as_of):
    rows = load(path, "4411 - Parked - Downpayment")
    hdr_i = find_row(rows, "Legal entity", col=0)
    data = [r for r in rows[hdr_i + 1:] if r[2] is not None]
    by_bu_origin = collections.defaultdict(float); cnt = collections.defaultdict(int)
    buckets = collections.defaultdict(float); tiers = collections.defaultdict(float); sup = collections.defaultdict(float)
    sup_bu = {}
    for r in data:
        bu = canon_bu(r[3]) or str(r[3]); origin = str(r[10]).strip(); amt = -num(r[14], 1e6)  # payables shown positive
        by_bu_origin[(bu, origin)] += amt; cnt[(bu, origin)] += 1
        d = r[9]
        if isinstance(d, (dt.datetime, dt.date)):
            dd = d.date() if isinstance(d, dt.datetime) else d
            days = (dd - as_of).days
            k = "Overdue" if days < 0 else ("Due ≤ 30d" if days <= 30 else ("Due 31–60d" if days <= 60 else "Due > 60d"))
        else: k = "No due date"
        buckets[(origin, k)] += amt
        tiers[str(r[6]).strip()] += amt
        s = str(r[5]).strip(); sup[s] += amt; sup_bu[s] = bu
    top = sorted(sup.items(), key=lambda x: -x[1])[:15]
    return {"by_bu_origin": {f"{k[0]}|{k[1]}": v for k, v in by_bu_origin.items()},
            "count": {f"{k[0]}|{k[1]}": v for k, v in cnt.items()},
            "buckets": {f"{k[0]}|{k[1]}": v for k, v in buckets.items()},
            "tiers": dict(tiers), "top_suppliers": [[s, v, sup_bu[s]] for s, v in top], "n_lines": len(data)}


# ----------------------------------------------------------------------------- AR baseline (raw AR line list, e.g. beginning of month)
def s2_code(label):
    """'C-5-a Signed by …' / 'C-5.a Signed by …' -> 'c-5.a'; 'PDR' -> 'pdr'."""
    l = str(label or "").strip()
    if len(l) >= 5 and l[1] == "-" and l[3] in "-.": return (l[:3] + "." + l[4]).lower()
    return l.lower()[:12]

def extract_ar_baseline(path):
    """Aggregate a raw AR line list with the AR dashboard rules (MMAD).
    Rules reproduce the dashboard snapshot: unbilled excludes 'J-6-a Unearned on LS'; advance excludes J-6-* and C-3-a lines;
    at risk = C-1-a/C-1-b; SS postponed = J-3-c; under JESA = other J-* codes; under OCP = C-* codes and PDR."""
    import warnings; warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["AR"] if "AR" in wb.sheetnames else wb.worksheets[0]
    it = ws.iter_rows(values_only=True); hdr = [str(h).strip() if h else "" for h in next(it)]
    I = {h: i for i, h in enumerate(hdr)}
    def g(r, name): return r[I[name]] if name in I else None
    kpi = collections.defaultdict(float); bu = collections.defaultdict(float); owner = collections.defaultdict(float)
    st2 = collections.defaultdict(float); unp = collections.defaultdict(float); aging = collections.defaultdict(float); labor = collections.defaultdict(float)
    for r in it:
        if not any(v is not None for v in r): continue
        s1 = str(g(r, "Status 1") or "").strip().lower()
        comp = "Advance" if s1.startswith("advance") else ("Unbilled" if s1.startswith("unbilled") else ("Unpaid" if s1.startswith("unpaid") else None))
        if comp is None: continue
        code = s2_code(g(r, "Status 2")); amt = num(g(r, "Total Amount HT"), 1e6)
        if comp == "Unbilled" and code == "j-6.a": continue
        if comp == "Advance" and (code.startswith("j-6") or code == "c-3.a"): continue
        b = canon_bu(g(r, "Business Unit")) or str(g(r, "Business Unit"))
        kpi[comp] += amt; bu[(b, comp)] += amt
        y = str(g(r, "Year") or "").strip(); aging[y if y[:2] == "20" else "No year (advance lines)"] += amt
        labor[(comp, "PSR" if str(g(r, "PSR/PROC") or "").strip().upper() == "PSR" else "PROC")] += amt
        if comp == "Unbilled":
            cat = "At risk" if code in ("c-1.a", "c-1.b") else ("SS postponed" if code == "j-3.c" else ("Under JESA" if code.startswith("j-") else "Under OCP"))
            owner[(b, cat)] += amt; st2[code] += amt
        if comp == "Unpaid":
            c = str(g(r, "OCP Comments") or "").strip().lower()
            lab = ("Autres / NON-OCP" if "hors ocp" in c else "OCP SUB" if "ocp sub" in c else "En ordonnancement" if c.startswith("en ordonnancement") else "Transmise tréso" if c.startswith("transmise") else
                   "Manque réception" if c.startswith("manque") else "Rejetée" if c.startswith("rejet") else "Autres / NON-OCP")
            unp[lab] += amt
    kpi["Total AR"] = kpi["Advance"] + kpi["Unbilled"] + kpi["Unpaid"]
    return {"kpi": dict(kpi), "bu": dict(bu), "owner": dict(owner), "status2": dict(st2), "unpaid": dict(unp), "aging": dict(aging), "labor": dict(labor)}

def apply_baseline(ar, base):
    """Overwrite the month_start columns of the dashboard extract with the baseline aggregation."""
    kmap = {"Advance Payment": "Advance", "Unbilled": "Unbilled", "Unpaid": "Unpaid", "Total AR": "Total AR"}
    for m in ar["kpi"]:
        if m[0] in kmap: m[1] = base["kpi"].get(kmap[m[0]], 0.0)
    for r in ar["ar_bu"]:
        r[2] = base["kpi"].get("x", 0.0) if False else (sum(v for (b, c), v in base["bu"].items() if b == r[0]) if r[1] == "Total AR" else base["bu"].get((r[0], r[1]), 0.0))
    for r in ar["unbilled_owner"]: r[2] = base["owner"].get((r[0], r[1]), 0.0)
    seen = set()
    for r in ar["unbilled_status"]:
        code = s2_code(r[1])
        if r[0] == "Other": continue
        r[2] = base["status2"].get(code, 0.0); seen.add(code)
    resid = sum(v for c, v in base["status2"].items() if c not in seen)
    for r in ar["unbilled_status"]:
        if r[0] == "Other": r[2] = resid
    for r in ar["unpaid_status"]: r[1] = base["unpaid"].get(r[0], 0.0)
    for r in ar["aging"]: r[1] = base["aging"].get(r[0], 0.0)
    return ar

# ----------------------------------------------------------------------------- Unrecognized revenue tracking (UR consolidation)
def extract_ur(path):
    out = {}
    su = load(path, "Summary UR")
    i = find_row(su, "Business Unit", col=0) + 1; rows = []; att = []
    while su[i][0] and str(su[i][0]).strip().lower() != "total":
        r = su[i]; b = canon_bu(r[0])
        if b:
            rows.append([b] + [num(r[c], 1e6) for c in range(1, 14)])
            att.append([b] + [num(r[c], 1e6) for c in range(16, 27)])
        i += 1
    out["table"] = rows; out["attachment"] = att
    prev = {}
    try:
        s1 = load(path, "Sheet1")
        hdr_i = find_row(s1, "Business Unit", col=1)
        # right block (including markup 6%) starts at the second 'Business Unit' header in the same row
        cols = [c for c, v in enumerate(s1[hdr_i]) if v and str(v).strip().lower() == "business unit"]
        bc = cols[-1]
        i = hdr_i + 1
        while s1[i][bc] and str(s1[i][bc]).strip().lower() != "total":
            b = canon_bu(s1[i][bc])
            if b: prev[b] = num(s1[i][bc + 1], 1e6)
            i += 1
    except Exception as e:
        print("WARN UR previous update:", e)
    out["prev"] = prev
    po = []
    try:
        sp = load(path, "Summary PO by BU"); i = find_row(sp, "Business Unit", col=0) + 1
        while sp[i][0] and str(sp[i][0]).strip().lower() != "total":
            b = canon_bu(sp[i][0])
            if b: po.append([b] + [num(sp[i][c], 1e6) for c in range(1, 5)])
            i += 1
    except Exception as e:
        print("WARN UR PO summary:", e)
    out["po"] = po
    return out

def apply_ur(ar, ur):
    cur = {r[0]: r[13] for r in ur["table"]}; tot_cur = sum(cur.values()); tot_prev = sum(ur["prev"].values()) if ur["prev"] else None
    ar["unrecognized_table"] = [r for r in ur["table"]]
    ar["unrecognized_bu"] = cur
    tot_ar = {m[0]: m for m in ar["kpi"]}["Total AR"]
    for m in ar["kpi"]:
        if m[0].startswith("Unrecognized") or m[0].startswith("Attachments"):
            m[3] = tot_cur
            if tot_prev is not None: m[2] = tot_prev
        if m[0].startswith("Total NEW AR"):
            m[3] = tot_ar[3] + tot_cur
            if tot_prev is not None: m[2] = tot_ar[2] + tot_prev
    return ar

# ----------------------------------------------------------------------------- Write workbook
HDR_FILL = PatternFill("solid", fgColor="1B2F4B"); HDR_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
NOTE_FONT = Font(italic=True, color="7A8794", name="Arial", size=9)

def write_sheet(wb, name, header, rows, note=None, widths=None):
    ws = wb.create_sheet(name)
    r0 = 1
    if note:
        ws.cell(row=1, column=1, value=note).font = NOTE_FONT; r0 = 2
    for j, h in enumerate(header, 1):
        c = ws.cell(row=r0, column=j, value=h); c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, row in enumerate(rows, r0 + 1):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=v)
            if isinstance(v, float): c.number_format = "#,##0.0;[Red]-#,##0.0"
            c.font = Font(name="Arial", size=10)
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = (widths[j - 1] if widths and j - 1 < len(widths) else 16)
    ws.freeze_panes = ws.cell(row=r0 + 1, column=2)
    return ws

def default_messages(ar, cash, lc, ap, meta):
    k = {m[0]: m for m in ar["kpi"]}
    tot = k["Total AR"]; unb = k["Unbilled"]; unp = k["Unpaid"]; adv = k["Advance Payment"]
    unrec = [m for m in ar["kpi"] if m[0].startswith("Unrecognized")][0]; new = [m for m in ar["kpi"] if m[0].startswith("Total NEW AR")][0]
    def d(x): return f"{x:+,.0f}".replace("-", "−")
    pct_unb = unb[3] / tot[3] * 100
    cm = {r[0]: r for r in cash["cash_monthly"]}
    net = cm["NET position (cash IN + cash OUT)"]; cin = cm["Total cash IN (procurement)"]; cout = cm["Total cash OUT (procurement)"]
    lc_open = sum(l["encours_mad"] for l in lc["lines"] if l["opened"] == "Yes"); lc_all = sum(l["encours_mad"] for l in lc["lines"])
    ap_tot = sum(v for v in ap["by_bu_origin"].values())
    ap_over = sum(v for kk, v in ap["buckets"].items() if kk.endswith("|Overdue"))
    coll = ar["collections_bu"]; oct_cur = sum(c[2] for c in coll); oct_pw = sum(c[1] for c in coll); ytd = sum(c[5] for c in coll)
    ms_lab = ("month-start (" + meta["ar_month_start_date"].strftime("%d %b") + ")") if meta.get("ar_month_start_date") else "month-start"
    return [
        ["Total AR " + ("down" if tot[3] < tot[2] else "up") + f" {abs(tot[3]-tot[2]):,.0f} MMAD vs last week and " + ("down" if tot[3] < tot[1] else "up") + f" {abs(tot[3]-tot[1]):,.0f} MMAD vs the {ms_lab} baseline",
         f"Total AR stands at {tot[3]:,.0f} MMAD ({d(tot[3]-tot[2])} WoW, {d(tot[3]-tot[1])} vs month start). Adding {unrec[3]:,.0f} MMAD of unrecognized revenue ({d(unrec[3]-unrec[2])} vs the previous BU update), gross exposure is {new[3]:,.0f} MMAD. "
         f"Unbilled is {unb[3]:,.0f} MMAD ({pct_unb:.0f}% of AR, {d(unb[3]-unb[2])} WoW); unpaid is {unp[3]:,.0f} MMAD ({d(unp[3]-unp[2])} WoW, {d(unp[3]-unp[1])} vs baseline)."],
        [f"Collections: {oct_cur:,.0f} MMAD received in October to date ({d(oct_cur-oct_pw)} WoW), {ytd:,.0f} MMAD YTD",
         f"October receipts moved from {oct_pw:,.0f} to {oct_cur:,.0f} MMAD this week. YTD payments received reach {ytd:,.0f} MMAD ({ytd/1000:,.2f} Bn MAD)."],
        [f"Procurement cash position: net {d(net[6])} MMAD Jul–Oct including October commitments",
         f"Cumulative procurement cash IN of {cin[6]:,.0f} MMAD vs cash OUT of {abs(cout[6]):,.0f} MMAD (Jul–Sep actuals, Oct actuals to {meta['payments_to']}, plus {abs(cout[5]):,.0f} MMAD of approved and pending October commitments). "
         f"October alone: {cin[4]:,.0f} MMAD in vs {abs(cout[4]):,.0f} MMAD paid, {abs(cout[5]):,.0f} MMAD still to pay."],
        [f"Payables: {ap_tot:,.0f} MMAD of supplier invoices in AP (4411 + parked + downpayments), {ap_over:,.0f} MMAD past due",
         f"{ap['n_lines']:,} open supplier lines. Parked invoices and 4411 open items dominate; see the AP slide for due-date buckets, tiers and top suppliers."],
        [f"Letters of credit: {lc_open:,.0f} MMAD outstanding on opened LCs, {lc_all:,.0f} MMAD including LCs still to be opened",
         f"October LC maturities of {cash['lc_oct_maturity_total']:,.0f} MMAD are counted as approved payments in the cash position. "
         f"{sum(l['part_2026'] for l in lc['lines']):,.0f} MMAD falls in 2026 and {sum(l['part_2027'] for l in lc['lines']):,.0f} MMAD in 2027 per the maturity schedule."],
    ]

def main():
    ap_ = argparse.ArgumentParser()
    ap_.add_argument("--ar", required=True); ap_.add_argument("--cash", required=True); ap_.add_argument("--lc", required=True); ap_.add_argument("--ap", required=True)
    ap_.add_argument("--as-of", required=True, help="deck date YYYY-MM-DD")
    ap_.add_argument("--ar-date", required=True, help="AR dashboard snapshot date YYYY-MM-DD")
    ap_.add_argument("--ar-prev-date", default=None, help="previous AR snapshot date (default: dashboard 'Previous baseline saved' - 1 day)")
    ap_.add_argument("--ar-baseline", default=None, help="raw AR line list used as month-start baseline (e.g. AR_Beg_of_the_month.xlsx)")
    ap_.add_argument("--ar-baseline-date", default=None, help="date label of the month-start baseline YYYY-MM-DD")
    ap_.add_argument("--ur", default=None, help="UR Tracking Consolidation workbook (unrecognized revenue)")
    ap_.add_argument("--collections-to", default=None); ap_.add_argument("--payments-to", default=None); ap_.add_argument("--lc-date", default=None)
    ap_.add_argument("--ap-label", default=None); ap_.add_argument("--week-label", default=None)
    ap_.add_argument("--monthly-seed", default=None, help="JSON file with monthly collections seed {month:[labor,nonlabor]}")
    ap_.add_argument("--out", default="inputs/weekly_inputs.xlsx")
    a = ap_.parse_args()

    as_of = dt.date.fromisoformat(a.as_of); ar_date = dt.date.fromisoformat(a.ar_date)
    ar = extract_ar(a.ar); cash = extract_cash(a.cash); lc = extract_lc(a.lc); ap = extract_ap(a.ap, as_of)
    ur = None
    if a.ar_baseline:
        base = extract_ar_baseline(a.ar_baseline); ar = apply_baseline(ar, base)
        print("baseline totals (MMAD):", {k: round(v, 1) for k, v in base["kpi"].items()})
        if a.ar_baseline_date: ar["month_start_date"] = dt.date.fromisoformat(a.ar_baseline_date)
    if a.ur:
        ur = extract_ur(a.ur); ar = apply_ur(ar, ur)
        print("unrecognized (MMAD): current", round(sum(r[13] for r in ur["table"]), 1), "previous", round(sum(ur["prev"].values()), 1))
    prev_date = dt.date.fromisoformat(a.ar_prev_date) if a.ar_prev_date else (ar["prev_baseline_date"] - dt.timedelta(days=1) if ar["prev_baseline_date"] else None)
    meta = {
        "as_of_date": as_of, "week_label": a.week_label or f"Week {as_of.isocalendar()[1]}",
        "ar_current_date": ar_date, "ar_prev_date": prev_date, "ar_month_start_date": ar["month_start_date"],
        "cash_data_as_of": cash["data_as_of"].date() if cash["data_as_of"] else None,
        "collections_to": a.collections_to or "19/10", "payments_to": a.payments_to or "15/10",
        "lc_date": a.lc_date or "19/10", "ap_label": a.ap_label or "Oct W3", "cash_basis": cash["basis"],
        "units": "MMAD (millions of MAD)",
    }

    wb = openpyxl.Workbook(); wb.remove(wb.active)
    write_sheet(wb, "META", ["key", "value"], [[k, (v.isoformat() if isinstance(v, dt.date) else v)] for k, v in meta.items()],
                note="Dates and labels used on the slides. Edit freely.", widths=[24, 60])
    write_sheet(wb, "KEY_MESSAGES", ["n", "title", "body"], [[i + 1, t, b] for i, (t, b) in enumerate(default_messages(ar, cash, lc, ap, meta))],
                note="Auto-drafted from the numbers; rewrite the wording as needed (max 5 messages).", widths=[4, 70, 120])
    write_sheet(wb, "AR_KPI", ["metric", "month_start", "prev_week", "current"], ar["kpi"],
                note="From AR dashboard 'Comparatives' – Key figures (MMAD). month_start = dashboard 'AR Baseline' tab.", widths=[44])
    ls = ar["labor_split"]
    write_sheet(wb, "AR_SPLITS", ["metric", "value"],
                [["Total AR labor", ls["total_labor"]], ["Total AR non-labor", ls["total_nonlabor"]], ["Unbilled labor", ls["unbilled_labor"]],
                 ["Unbilled non-labor", ls["unbilled_nonlabor"]], ["Unpaid labor", ls["unpaid_labor"]], ["Unpaid non-labor", ls["unpaid_nonlabor"]]] +
                [[f"AR {k}", v] for k, v in ar["ocp_split"].items()], note="Current week only (AR dashboard 'Executive Summary').", widths=[30])
    write_sheet(wb, "AR_BU", ["bu", "component", "month_start", "prev_week", "current"],
                sorted(ar["ar_bu"], key=lambda r: (BU_ORDER.index(r[0]) if r[0] in BU_ORDER else 99, ["Advance", "Unbilled", "Unpaid", "Total AR"].index(r[1]))) +
                [[b, "Unrecognized", None, None, v] for b, v in ar["unrecognized_bu"].items()],
                note="AR by business unit and component (MMAD). Unrecognized revenue = current only.", widths=[28, 16])
    write_sheet(wb, "AR_UNBILLED_OWNER", ["bu", "category", "month_start", "prev_week", "current"], ar["unbilled_owner"],
                note="Unbilled split by ownership: At risk / Under JESA / Under OCP / Service sheets postponed.", widths=[28, 16])
    write_sheet(wb, "AR_UNBILLED_STATUS", ["group", "status", "month_start", "prev_week", "current"], ar["unbilled_status"],
                note="Unbilled by detailed status (new coded scheme). month_start may be n/a (0) for statuses created after the baseline.", widths=[10, 50])
    write_sheet(wb, "AR_UNPAID_STATUS", ["status", "month_start", "prev_week", "current"], ar["unpaid_status"],
                note="Unpaid by OCP processing status.", widths=[28])
    write_sheet(wb, "AR_AGING", ["year", "month_start", "prev_week", "current"], ar["aging"], note="Total AR by year of origin.", widths=[24])
    write_sheet(wb, "UNRECOGNIZED", ["bu", "eq_pending_approval", "eq_signed_by_client", "eq_pending_submission", "eq_under_cn", "eq_at_risk",
                                     "sub_pending_approval", "sub_signed_by_client", "sub_pending_submission", "sub_under_cn", "sub_at_risk",
                                     "total_suppliers", "lines_to_adjust", "grand_total", "prev_week_total"],
                [r[:14] + [(ur["prev"].get(r[0]) if ur else None)] for r in ar["unrecognized_table"]],
                note="Unrecognized suppliers (equipment + subcontracts) by status, MMAD incl. markup, current week; prev_week_total = previous BU update (incl. markup).", widths=[28] + [14] * 14)
    if ur:
        write_sheet(wb, "UR_ATTACHMENT_STATUS", ["bu", "eq_under_jesa_approval", "eq_under_jesa_preparation", "eq_under_client_approval", "eq_signed_by_client", "eq_accrual_posted",
                                                 "sub_under_jesa_approval", "sub_under_jesa_preparation", "sub_under_client_approval", "sub_signed_by_client", "sub_accrual_posted", "total"],
                    ur["attachment"], note="Status of JESA attachments to client (UR consolidation, MMAD incl. markup).", widths=[28] + [14] * 11)
        write_sheet(wb, "UR_PO_SUMMARY", ["bu", "contract_value", "recognized", "unrecognized", "remaining"], ur["po"],
                    note="Supplier PO consolidation by BU (MMAD): contract value, recognized (SAP GR), unrecognized, remaining.", widths=[28, 16, 16, 16, 16])
    write_sheet(wb, "COLLECTIONS_BU", ["bu", "month_prev_week", "month_current", "prior_month", "ytd_prev_week", "ytd_current", "invoices_ytd"],
                sorted(ar["collections_bu"], key=lambda r: -r[5]), note="Payments received (collections) by BU: current month to date vs last week, prior full month, YTD.", widths=[28])
    # monthly collections series (seed + current month override)
    seed = {"Jan": [757, 1429], "Feb": [542, 1126], "Mar": [179, 306], "Apr": [59, 132], "May": [164, 299], "Jun": [91, 90],
            "Jul": [116, 695], "Aug": [136, 719], "Sep": [67, 341], "Oct": [28, 82], "Nov": [0, 0], "Dec": [0, 0]}
    if a.monthly_seed:
        seed.update(json.load(open(a.monthly_seed)))
    cur_month = as_of.strftime("%b"); cur_tot = sum(c[2] for c in ar["collections_bu"])
    l0, n0 = seed[cur_month]
    if (l0 + n0) > 0: seed[cur_month] = [cur_tot * l0 / (l0 + n0), cur_tot * n0 / (l0 + n0)]
    else: seed[cur_month] = [0.0, cur_tot]
    write_sheet(wb, "COLLECTIONS_MONTHLY", ["month", "labor", "non_labor"], [[m, float(v[0]), float(v[1])] for m, v in seed.items()],
                note="Monthly collections (MMAD), labor (PSR) vs non-labor (procurement). Jan–Sep seeded from the 14-Oct deck; current month total overwritten from the AR dashboard (L/NL pro-rated). Update by hand if you have the exact split.", widths=[10])
    write_sheet(wb, "CASH_MONTHLY", ["line", "jul", "aug", "sep", "oct_actual", "oct_ongoing", "cumulative"], cash["cash_monthly"],
                note="Cash flow workbook 'EXEC Summary' section 1 & 3 (MMAD): cash in positive, cash out negative. oct_ongoing = approved + pending October commitments.", widths=[48])
    write_sheet(wb, "CASH_BU", ["bu", "cash_in", "cash_out_suppliers", "cash_out_lc", "cash_out_customs", "net_actuals", "oct_approved", "net_after_approved", "oct_pending", "net_all_ongoing"],
                cash["cash_bu"], note="Procurement net balance by BU, Jul–Oct cumulative (EXEC Summary section 5).", widths=[28])
    bu_oct = []
    for b in BU_ORDER:
        if b in cash["bu_cash_in"] or b in cash["bu_cash_out"]:
            ci = cash["bu_cash_in"].get(b, [0] * 6); co = cash["bu_cash_out"].get(b, [0] * 6); lcv = cash["bu_lc"].get(b, [0] * 6); cu = cash["bu_customs"].get(b, [0] * 6)
            bu_oct.append([b, ci[3], co[3], lcv[3], cu[3], co[4], lcv[4]])
    write_sheet(wb, "CASH_BU_OCT", ["bu", "oct_cash_in", "oct_paid_suppliers", "oct_paid_lc", "oct_paid_customs", "oct_ongoing_suppliers", "oct_ongoing_lc"], bu_oct,
                note="October month only: collections to date and payments to date by BU, plus October open commitments (approved + pending).", widths=[28])
    cm = {r[0]: r for r in cash["cash_monthly"]}
    ongoing = [["Paid in October to date (suppliers + LC + customs)", None, abs(cm["Total cash OUT (procurement)"][4])],
               ["Approved – October LC maturities", None, cash["lc_oct_maturity_total"]],
               ["Approved – supplier invoices", None, abs(cm["Cash OUT suppliers (by tier)"][5]) - 0.0],
               ["Pending payment approval", None, None]]
    # approved + pending from section 5 totals
    tot_appr = sum(r[6] for r in cash["cash_bu"]); tot_pend = sum(r[8] for r in cash["cash_bu"])
    ongoing[2][2] = abs(tot_appr) - cash["lc_oct_maturity_total"]; ongoing[3][2] = abs(tot_pend)
    ongoing += [[f"Tracking log – {r[0]}", r[1], r[2]] for r in cash["approved_status"]]
    ongoing += [["Priority list – total approved", None, cash["priority_total"]], ["Priority list – carried in October ongoing", None, cash["priority_carried"]]]
    ongoing += [[f"Priority list – {k}", None, v] for k, v in cash["priority_by_class"].items()]
    write_sheet(wb, "ONGOING", ["bucket", "lines", "amount"], ongoing, note="October ongoing payments position (MMAD, positive = to pay).", widths=[60, 8])
    write_sheet(wb, "AP_BU", ["bu", "downpayment_prev", "downpayment_cur", "parked_prev", "parked_cur", "inv4411_prev", "inv4411_cur", "lc_prev", "lc_cur"], ar["ap_bu"],
                note="Supplier payables & LC by BU from the AR dashboard (MMAD), previous week vs current.", widths=[28])
    write_sheet(wb, "AP_BUCKETS", ["origin", "bucket", "amount"], [[k.split("|")[0], k.split("|")[1], v] for k, v in sorted(ap["buckets"].items())],
                note=f"Supplier invoice file ({meta['ap_label']}): open payables by origin and due-date bucket vs {as_of.isoformat()} (MMAD, positive = payable).", widths=[20, 16])
    write_sheet(wb, "AP_BU_ORIGIN", ["bu", "origin", "amount", "lines"], [[k.split("|")[0], k.split("|")[1], v, ap["count"][k]] for k, v in sorted(ap["by_bu_origin"].items())],
                note="Supplier invoice file: open payables by BU and origin.", widths=[28, 20])
    write_sheet(wb, "AP_TIERS", ["tier", "amount"], sorted([[k, v] for k, v in ap["tiers"].items()], key=lambda r: -r[1]), note="Supplier invoice file: open payables by supplier tier.", widths=[24])
    write_sheet(wb, "AP_TOP_SUPPLIERS", ["supplier", "amount", "bu"], ap["top_suppliers"], note="Top 15 suppliers by open payables (supplier invoice file).", widths=[36, 12, 28])
    write_sheet(wb, "LC_LINES", ["project", "bu", "supplier", "lc_ref", "expiry", "ccy", "face_value_fx", "encours_fx", "encours_mad", "opened", "action", "part_2026", "part_2027"] + lc["months"],
                [[l["project"], l["bu"], l["supplier"], l["ref"], l["expiry"], l["ccy"], l["face_fx"], l["encours_fx"], l["encours_mad"], l["opened"], l["action"], l["part_2026"], l["part_2027"]] + [l["months"][m] for m in lc["months"]] for l in lc["lines"]],
                note=f"LC situation {meta['lc_date']}: one line per LC (encours in MMAD; maturity schedule by month in MMAD). opened = LC reference exists.", widths=[14, 24, 30, 18, 12, 6, 14, 14, 12, 8, 36, 10, 10])
    wb.save(a.out)
    json.dump({"meta": {k: (v.isoformat() if isinstance(v, dt.date) else v) for k, v in meta.items()}}, open(a.out.replace(".xlsx", "_meta.json"), "w"), indent=1, default=str)
    print("written", a.out)

if __name__ == "__main__":
    main()
