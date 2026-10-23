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
import os, sys; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
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
    for lab, anchor in [("PSR provision – AM ring-fence", "AM PSR (Labor)"), ("PSR provision – W&E reserve", "Water & Environnement — PSR reserve"),
                        ("Overhead payments (FYI, not deducted)", "FYI — Overhead")]:
        try:
            vals, _ = series(anchor, start=s3); monthly.append([lab] + vals)
        except Exception as e:
            print("WARN PSR line:", lab, e)
    out["cash_monthly"] = monthly
    # PSR (labor) collections by BU (section 3)
    psr_bu = []; i = s3 + 1
    while i < len(ex) and not str(ex[i][2] or "").strip().lower().startswith("total psr"):
        b = canon_bu(ex[i][2]) if ex[i][2] else None
        if b: psr_bu.append([b] + [num(ex[i][c], 1e6) for c in range(3, 9)])
        i += 1
    out["psr_bu"] = psr_bu

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

# ----------------------------------------------------------------------------- DSO calculation workbook (WIP_DSO_calc.xlsx)
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
def extract_dso_calc(path, as_of):
    """12-month trailing revenue per BU ('12 Months Trailing' sheet) + DSO history ('Trend' sheet) + the workbook's own DSO table."""
    out = {"rev12": {}, "window": [], "trend": {}, "trend_months": [], "table": {}}
    wbk = openpyxl.load_workbook(path, data_only=True, read_only=True)
    def load(_, sheet): return [list(r) for r in wbk[sheet].iter_rows(values_only=True)]
    tr = load(path, "12 Months Trailing")
    hdr = [str(h).replace("\n", " ").strip().lower() if h else "" for h in tr[0]]
    i_sum = next((i for i, h in enumerate(hdr) if h.startswith("12 month")), None)
    i_bu = next((i for i, h in enumerate(hdr) if h == "ar bu"), None)
    if i_sum is None or i_bu is None: raise ValueError("'12 Months Trailing': columns '12 Month trailing' / 'AR BU' not found")
    out["window"] = [v.strftime("%b-%y") for v in tr[1] if isinstance(v, (dt.datetime, dt.date))]
    rev = collections.defaultdict(float)
    for r in tr[2:]:
        b = canon_bu(r[i_bu]) if r[i_bu] else None
        if b: rev[b] += num(r[i_sum], 1e6)
    out["rev12"] = dict(rev)
    try:
        td = load(path, "Trend"); months = [str(m).strip() for m in td[0][1:] if m]
        cur_name = months[0]
        # history columns: all months after the first (current) one; year inferred backwards from the as-of month
        y = as_of.year; labels = []
        for m in months[1:]:
            mi = MONTHS.index(m) + 1
            yy = as_of.year if mi < as_of.month else as_of.year - 1
            labels.append(f"{m[:3]}-{str(yy)[2:]}")
        out["trend_months"] = labels
        for r in td[1:]:
            b = canon_bu(r[0]) if r[0] else None
            if b: out["trend"][b] = [num(v, 1.0) if isinstance(v, (int, float)) else None for v in r[2:2 + len(labels)]]
    except Exception as e:
        print("WARN DSO trend:", e)
    try:
        tb_ = load(path, "DSO Table")
        for r in tb_[1:]:
            b = canon_bu(r[0]) if r[0] else ("JESA Group" if r[0] and str(r[0]).strip().lower() == "total" else None)
            if b and isinstance(r[3], (int, float)): out["table"][b] = float(r[3])
    except Exception as e:
        print("WARN DSO table:", e)
    return out

def compute_dso(dc, ar_bu_total):
    """DSO (days) = total AR of the BU / 12-month trailing revenue x 365; group = totals over the BUs with revenue."""
    dso = {}; tot_ar = 0.0; tot_rev = 0.0
    for b, rev in dc["rev12"].items():
        ar_v = ar_bu_total.get(b, 0.0)
        if rev > 0.05:
            dso[b] = {"current": round(ar_v / rev * 365), "prior": None, "ar": ar_v, "rev12": rev}; tot_ar += ar_v; tot_rev += rev
    if tot_rev: dso["JESA Group"] = {"current": round(tot_ar / tot_rev * 365), "prior": None, "ar": tot_ar, "rev12": tot_rev}
    for b, hist in dc["trend"].items():
        if b in dso and hist: dso[b]["prior"] = hist[-1]
    return dso

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
    def d(x): return "0" if abs(x) < 0.5 else f"{x:+,.0f}".replace("-", "−")
    def dd(m): return d(round(m[3]) - round(m[2]))
    def dm(m): return d(round(m[3]) - round(m[1])) if m[1] else "n/a"
    ms_lab = meta["ar_month_start_date"].strftime("%d %b") if meta.get("ar_month_start_date") else "month start"
    pw_lab = meta["ar_prev_date"].strftime("%d %b"); cw_lab = meta["ar_current_date"].strftime("%d %b")
    cm = {r[0]: r for r in cash["cash_monthly"]}
    cin = cm["Total cash IN (procurement)"]; cout = cm["Total cash OUT (procurement)"]
    cb = cash["cash_bu"]; net_act = sum(r[5] for r in cb); appr = -sum(r[6] for r in cb); net_appr = sum(r[7] for r in cb); pend = -sum(r[8] for r in cb); net_all = sum(r[9] for r in cb)
    out_act = -(cout[1] + cout[2] + cout[3] + cout[4])
    lc_open = sum(l["encours_mad"] for l in lc["lines"] if l["opened"] == "Yes"); lc_all = sum(l["encours_mad"] for l in lc["lines"])
    mt = collections.defaultdict(float)
    for l in lc["lines"]:
        for m, v in l["months"].items(): mt[m[:4]] += v
    lc_oct = sum(l["months"].get(lc["months"][0], 0) for l in lc["lines"]) if lc["months"] else 0
    ap_tot = sum(v for v in ap["by_bu_origin"].values()); ap_over = sum(v for kk, v in ap["buckets"].items() if kk.endswith("|Overdue"))
    o = {kk.split("|")[1]: 0 for kk in ap["by_bu_origin"]}
    for kk, v in ap["by_bu_origin"].items(): o[kk.split("|")[1]] += v
    coll = ar["collections_bu"]; oct_cur = sum(c[2] for c in coll); oct_pw = sum(c[1] for c in coll); ytd = sum(c[5] for c in coll); dep = sum(c[7] for c in coll)
    month = meta["as_of_date"].strftime("%B")
    ju = ar.get("jesa_unbilled_tot"); nonocp = ar.get("nonocp_tot")
    return [
        [f"Total AR {tot[3]:,.0f} MMAD: {dd(tot)} vs {pw_lab}, {dm(tot)} vs {ms_lab}",
         f"Unbilled {unb[3]:,.0f} ({dd(unb)} WoW), unpaid {unp[3]:,.0f} ({dd(unp)} WoW), advance {adv[3]:,.0f} ({dd(adv)} WoW)."
         + (f" Unbilled under JESA responsibility {ju[0]:,.0f} MMAD (PSR {ju[1]:,.0f} / Procurement {ju[2]:,.0f}); non-OCP receivables {nonocp:,.0f} MMAD." if ju else "")],
        [f"Unrecognized revenue {unrec[3]:,.0f} MMAD ({dd(unrec)} WoW, {dm(unrec)} vs {ms_lab}); gross exposure {new[3]:,.0f} MMAD ({dd(new)} WoW)",
         f"Executed supplier and equipment scope not yet booked as AR, including the 6% JESA cost + fee."],
        [f"Invoices paid in {month} to {cw_lab}: {oct_cur:,.0f} MMAD ({d(round(oct_cur)-round(oct_pw))} vs {pw_lab}); {ytd:,.0f} MMAD year to date",
         f"Invoices deposited in {month}: {dep:,.0f} MMAD."],
        [f"Procurement cash Jul–{month[:3]}: {d(net_act)} MMAD on actuals, {d(net_appr)} after approved payments, {d(net_all)} after pending payments",
         f"Cash IN {cin[6]:,.0f} MMAD vs cash OUT {out_act:,.0f} MMAD paid to {meta['payments_to']} gives {d(net_act)} MMAD. "
         f"Approved payments of {appr:,.0f} MMAD (incl. {cash['lc_oct_maturity_total']:,.0f} of {month} LC maturities) bring the position to {d(net_appr)} MMAD; "
         f"pending payments of {pend:,.0f} MMAD bring it to {d(net_all)} MMAD."],
        [f"Supplier payables {ap_tot:,.0f} MMAD, of which {ap_over:,.0f} MMAD past due",
         f"Parked invoices {o.get('Parked Invoice', 0):,.0f}, 4411 open items {o.get('Invoices - 4411', 0):,.0f}, downpayments {o.get('Downpayment', 0):,.0f} MMAD ({ap['n_lines']:,} open lines)."],
        [f"Letters of credit {lc_all:,.0f} MMAD: {lc_open:,.0f} on opened LCs, {lc_all - lc_open:,.0f} still to be opened",
         f"{lc_oct:,.0f} MMAD matures in {month}; {mt.get('2026', 0):,.0f} MMAD in 2026 and {mt.get('2027', 0):,.0f} MMAD in 2027."],
    ]

def prev_from_dashboard(d):
    """Previous-week structures (same shape as ar_engine.aggregate) from the processed AR dashboard's 'current' columns."""
    k = {m[0]: m for m in d["kpi"]}
    P = {"kpi": {"Advance": k["Advance Payment"][3], "Unbilled": k["Unbilled"][3], "Unpaid": k["Unpaid"][3], "Total AR": k["Total AR"][3]},
         "bu": {(r[0], r[1]): r[4] for r in d["ar_bu"]}, "owner": {(r[0], r[1]): r[4] for r in d["unbilled_owner"]},
         "status_total": {s2_code(r[1]): r[4] for r in d["unbilled_status"] if r[0] != "Other"},
         "unpaid_phase_total": {}, "year": {r[0]: r[3] for r in d["aging"]},
         "ap": {r[0]: {"downpayment": r[2], "parked": r[4], "inv4411": r[6], "lc": r[8]} for r in d["ap_bu"]}}
    m = {"En ordonnancement": "Ordonnancement", "Transmise tréso": "Treasury", "Manque réception": "Missing reception", "Rejetée": "Rejected", "OCP SUB": "OCP subs", "Autres / NON-OCP": "Other"}
    for r in d["unpaid_status"]: P["unpaid_phase_total"][m.get(r[0], "Other")] = P["unpaid_phase_total"].get(m.get(r[0], "Other"), 0) + r[3]
    return P

def prev_from_inputs(path):
    """Previous-week structures from last week's weekly_inputs.xlsx (the standard weekly roll-forward)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    def recs(name):
        ws = wb[name]; rows = [list(r) for r in ws.iter_rows(values_only=True)]
        h = 1 if rows and isinstance(rows[0][0], str) and all(v is None for v in rows[0][1:]) else 0
        hdr = [str(c) for c in rows[h]]
        return [dict(zip(hdr, r)) for r in rows[h + 1:] if any(v is not None for v in r)]
    kpi = {r["metric"]: r["current"] for r in recs("AR_KPI")}
    P = {"kpi": {"Advance": kpi.get("Advance Payment", 0), "Unbilled": kpi.get("Unbilled", 0), "Unpaid": kpi.get("Unpaid", 0), "Total AR": kpi.get("Total AR", 0)},
         "bu": {(r["bu"], r["component"]): r["current"] or 0 for r in recs("AR_BU")},
         "owner": {(r["bu"], r["category"]): r["current"] or 0 for r in recs("AR_UNBILLED_OWNER")},
         "status_total": {s2_code(r["status"]): r["current"] or 0 for r in recs("AR_UNBILLED_STATUS") if r["group"] != "Other"},
         "unpaid_phase_total": {r["status"]: r["current"] or 0 for r in recs("AR_UNPAID_STATUS")},
         "year": {str(r["year"]): r["current"] or 0 for r in recs("AR_AGING")},
         "ap": {r["bu"]: {"downpayment": r["downpayment_cur"], "parked": r["parked_cur"], "inv4411": r["inv4411_cur"], "lc": r["lc_cur"]} for r in recs("AP_BU")}}
    if "AR_BU_DETAIL" in wb.sheetnames:
        P["detail"] = {(r["bu"], r["section"], r["item"]): r["current"] or 0 for r in recs("AR_BU_DETAIL")}
    if "UNRECOGNIZED" in wb.sheetnames:
        P["ur_prev"] = {r["bu"]: r["grand_total"] or 0 for r in recs("UNRECOGNIZED")}
        P["ur_ms"] = {r["bu"]: r["month_start_total"] for r in recs("UNRECOGNIZED") if r.get("month_start_total") is not None}
    if "DSO" in wb.sheetnames:
        P["dso"] = {r["bu"]: {"current": r["dso_current"], "prior": r["dso_prior_month"]} for r in recs("DSO")}
    if "NONOCP_BU" in wb.sheetnames:
        P["nonocp"] = {(r["bu"], r["component"]): r["current"] or 0 for r in recs("NONOCP_BU")}
    if "JESA_UNBILLED_BU" in wb.sheetnames:
        P["jesa_unbilled"] = {(r["bu"], r["split"]): r["current"] or 0 for r in recs("JESA_UNBILLED_BU")}
    return P

PHASES = ["Rejected", "Treasury", "OCP subs", "Ordonnancement", "Missing reception", "Other"]
def phase_total(agg, b=None):
    out = collections.defaultdict(float)
    for (bb, ph), v in agg["unpaid_phase"].items():
        if b and bb != b: continue
        out[ph if ph in PHASES else "Other"] += v
    return out

def main():
    ap_ = argparse.ArgumentParser()
    ap_.add_argument("--ar", required=True, help="raw AR line list of the current week (sheet 'AR')")
    ap_.add_argument("--ar-prev", default=None, help="raw AR line list of the previous week (optional)")
    ap_.add_argument("--ar-dashboard", default=None, help="processed AR dashboard (Executive Summary + Comparatives) whose 'current' columns are last week (transition only)")
    ap_.add_argument("--prev-inputs", default=None, help="last week's weekly_inputs.xlsx (standard source of previous-week values)")
    ap_.add_argument("--ar-baseline", default=None, help="raw AR line list used as month-start baseline")
    ap_.add_argument("--ar-baseline-date", default=None, help="date label of the month-start baseline YYYY-MM-DD")
    ap_.add_argument("--cash", required=True); ap_.add_argument("--lc", required=True); ap_.add_argument("--ap", required=True)
    ap_.add_argument("--ur", default=None, help="UR Tracking Consolidation workbook (unrecognized revenue)")
    ap_.add_argument("--ur-ms", default=None, help="JSON {bu: MMAD} with month-start unrecognized revenue by BU")
    ap_.add_argument("--dso", default=None, help="JSON {bu: {current, prior}} DSO in days (manual override, e.g. from the BU meeting report)")
    ap_.add_argument("--dso-calc", default=None, help="WIP_DSO_calc.xlsx: 12-month trailing revenue per BU and DSO trend; DSO is then computed from the AR extract")
    ap_.add_argument("--ur-prev", default=None, help="JSON {bu: MMAD} with last week's unrecognized revenue by BU (overrides the UR file's 'previous BU update' and the prev inputs)")
    ap_.add_argument("--as-of", required=True, help="deck date YYYY-MM-DD")
    ap_.add_argument("--ar-date", required=True, help="AR snapshot date YYYY-MM-DD")
    ap_.add_argument("--ar-prev-date", required=True, help="previous AR snapshot date YYYY-MM-DD")
    ap_.add_argument("--collections-to", default=None); ap_.add_argument("--payments-to", default=None); ap_.add_argument("--lc-date", default=None)
    ap_.add_argument("--ap-label", default=None); ap_.add_argument("--week-label", default=None)
    ap_.add_argument("--out", default="inputs/weekly_inputs.xlsx")
    a = ap_.parse_args()

    as_of = dt.date.fromisoformat(a.as_of); ar_date = dt.date.fromisoformat(a.ar_date); prev_date = dt.date.fromisoformat(a.ar_prev_date)
    from ar_engine import aggregate, bus_in, status_label, owner_of, BUCKETS
    cur = aggregate(a.ar, ar_date, prev_cutoff=prev_date)
    print("current AR (MMAD):", {k: round(v, 1) for k, v in cur["kpi"].items()})
    if a.ar_prev:
        P = aggregate(a.ar_prev, prev_date); P["detail"] = None
        P["ap"] = {}; P["unpaid_phase_total"] = dict(phase_total(P)); P["status_total"] = P["status_total"]
        P["detail_agg"] = P
    elif a.prev_inputs: P = prev_from_inputs(a.prev_inputs)
    elif a.ar_dashboard: P = prev_from_dashboard(extract_ar(a.ar_dashboard))
    else: raise SystemExit("need one of --ar-prev / --prev-inputs / --ar-dashboard for the previous week")
    print("previous AR (MMAD):", {k: round(v, 1) for k, v in P["kpi"].items()})
    MS = aggregate(a.ar_baseline, dt.date.fromisoformat(a.ar_baseline_date) if a.ar_baseline_date else ar_date) if a.ar_baseline else None
    if MS: print("month-start AR (MMAD):", {k: round(v, 1) for k, v in MS["kpi"].items()})
    cash = extract_cash(a.cash); lc = extract_lc(a.lc); ap = extract_ap(a.ap, as_of)
    ur = extract_ur(a.ur) if a.ur else None
    if ur and a.ur_prev:
        ur["prev"] = {canon_bu(k) or k: float(v) for k, v in json.load(open(a.ur_prev)).items()}; ur["prev_source"] = "last week's deck"
    elif ur and P.get("ur_prev"):
        ur["prev"] = P["ur_prev"]; ur["prev_source"] = "last week's inputs"
    elif ur: ur["prev_source"] = "UR file previous BU update (incl. markup)"
    ur_ms = {canon_bu(k) or k: float(v) for k, v in json.load(open(a.ur_ms)).items()} if a.ur_ms else (P.get("ur_ms") or {})
    dso = {}; dso_trend = None
    if a.dso_calc:
        dc = extract_dso_calc(a.dso_calc, as_of)
        ar_bu_total = collections.defaultdict(float)
        for (b, comp), v in cur["bu"].items(): ar_bu_total[b] += v
        dso = compute_dso(dc, ar_bu_total); dso_trend = dc
        for b, v in dso.items():
            ref = dc["table"].get(b)
            if ref is not None and abs(ref - v["current"]) >= 2: print(f"NOTE DSO {b}: computed {v['current']} days vs {ref:.0f} in the DSO workbook table")
        print("DSO (days):", {b: v["current"] for b, v in dso.items()})
    elif a.dso:
        for k, v in json.load(open(a.dso)).items():
            if k.startswith("_"): continue
            dso[canon_bu(k) or k] = v
    elif P.get("dso"): dso = P["dso"]
    ms_date = dt.date.fromisoformat(a.ar_baseline_date) if a.ar_baseline_date else None

    # ------------------------------------------------------------------ assemble the 'ar' structure used by the writers
    def msv(d, k): return (d.get(k, 0.0) if MS else 0.0)
    ur_cur = sum(r[13] for r in ur["table"]) if ur else 0.0; ur_prev = sum(ur["prev"].values()) if (ur and ur["prev"]) else 0.0; ur_ms_tot = sum(ur_ms.values()) if ur_ms else 0.0
    ap_cur_bu = collections.defaultdict(lambda: {"downpayment": 0.0, "parked": 0.0, "inv4411": 0.0, "lc": 0.0})
    for k, v in ap["by_bu_origin"].items():
        b, o = k.split("|"); key = {"Downpayment": "downpayment", "Parked Invoice": "parked", "Invoices - 4411": "inv4411"}.get(o)
        if key: ap_cur_bu[b][key] += v
    for l in lc["lines"]:
        if l["opened"] == "Yes": ap_cur_bu[l["bu"]]["lc"] += l["encours_mad"]
    ap_prev_bu = P.get("ap", {})
    def ap_tot(d, k): return sum(v[k] for v in d.values())
    kpi = [["Advance Payment", msv(MS["kpi"] if MS else {}, "Advance"), P["kpi"]["Advance"], cur["kpi"]["Advance"]],
           ["Unbilled", msv(MS["kpi"] if MS else {}, "Unbilled"), P["kpi"]["Unbilled"], cur["kpi"]["Unbilled"]],
           ["Unpaid", msv(MS["kpi"] if MS else {}, "Unpaid"), P["kpi"]["Unpaid"], cur["kpi"]["Unpaid"]],
           ["Total AR", msv(MS["kpi"] if MS else {}, "Total AR"), P["kpi"]["Total AR"], cur["kpi"]["Total AR"]],
           ["Unrecognized Revenue (incl. markup)", ur_ms_tot, ur_prev, ur_cur],
           ["Total NEW AR", (msv(MS["kpi"] if MS else {}, "Total AR") + ur_ms_tot) if (MS and ur_ms) else 0, P["kpi"]["Total AR"] + ur_prev, cur["kpi"]["Total AR"] + ur_cur],
           ["AP Downpayment", 0, ap_tot(ap_prev_bu, "downpayment"), ap_tot(ap_cur_bu, "downpayment")],
           ["AP Parked Invoices", 0, ap_tot(ap_prev_bu, "parked"), ap_tot(ap_cur_bu, "parked")],
           ["AP Invoices 4411", 0, ap_tot(ap_prev_bu, "inv4411"), ap_tot(ap_cur_bu, "inv4411")],
           ["LC", 0, ap_tot(ap_prev_bu, "lc"), ap_tot(ap_cur_bu, "lc")]]
    kpi.append(["Total Payables incl. LC", 0, sum(r[2] for r in kpi[6:10]), sum(r[3] for r in kpi[6:10])])
    bus = [b for b in bus_in(cur) if b != "Other"]
    ar = {"kpi": kpi, "month_start_date": ms_date, "prev_baseline_date": None}
    ar["ar_bu"] = []
    for b in bus:
        for comp in ("Advance", "Unbilled", "Unpaid"):
            ar["ar_bu"].append([b, comp, msv(MS["bu"] if MS else {}, (b, comp)), P["bu"].get((b, comp), 0.0), cur["bu"].get((b, comp), 0.0)])
        ar["ar_bu"].append([b, "Total AR", sum(msv(MS["bu"] if MS else {}, (b, c)) for c in ("Advance", "Unbilled", "Unpaid")),
                            sum(P["bu"].get((b, c), 0.0) for c in ("Advance", "Unbilled", "Unpaid")), sum(cur["bu"].get((b, c), 0.0) for c in ("Advance", "Unbilled", "Unpaid"))])
    ar["unbilled_owner"] = [[b, cat, msv(MS["owner"] if MS else {}, (b, cat)), P["owner"].get((b, cat), 0.0), cur["owner"].get((b, cat), 0.0)]
                            for b in bus for cat in ("At risk", "Under JESA", "Under OCP", "SS postponed")]
    codes = sorted(set(cur["status_total"]) | set(P["status_total"]) | (set(MS["status_total"]) if MS else set()), key=lambda c: -cur["status_total"].get(c, 0))
    grp = {"At risk": "At Risk", "Under JESA": "JESA", "Under OCP": "OCP", "SS postponed": "OCP"}
    ar["unbilled_status"] = [[grp[owner_of(c)], status_label(c), msv(MS["status_total"] if MS else {}, c), P["status_total"].get(c, 0.0), cur["status_total"].get(c, 0.0)] for c in codes]
    cur_ph = phase_total(cur); ms_ph = phase_total(MS) if MS else {}
    ar["unpaid_status"] = [[ph, ms_ph.get(ph, 0.0), P["unpaid_phase_total"].get(ph, 0.0), cur_ph.get(ph, 0.0)] for ph in PHASES]
    years = sorted(set(cur["year"]) | set(P["year"]) | (set(MS["year"]) if MS else set()), key=lambda y: (y[:2] != "20", y))
    ar["aging"] = [[y, msv(MS["year"] if MS else {}, y), P["year"].get(y, 0.0), cur["year"].get(y, 0.0)] for y in years]
    ar["labor_split"] = {"total_labor": sum(v for (c, l), v in cur["labor"].items() if l == "labor"), "total_nonlabor": sum(v for (c, l), v in cur["labor"].items() if l == "non_labor"),
                         "unbilled_labor": cur["labor"].get(("Unbilled", "labor"), 0.0), "unbilled_nonlabor": cur["labor"].get(("Unbilled", "non_labor"), 0.0),
                         "unpaid_labor": cur["labor"].get(("Unpaid", "labor"), 0.0), "unpaid_nonlabor": cur["labor"].get(("Unpaid", "non_labor"), 0.0)}
    ar["ocp_split"] = {"OCP": cur["ocp"].get("OCP", 0.0), "NON OCP": cur["ocp"].get("NON OCP", 0.0)}
    ar["unrecognized_bu"] = ({r[0]: r[13] for r in ur["table"]} if ur else {})
    ar["unrecognized_table"] = (ur["table"] if ur else [])
    ar["collections_bu"] = []
    for b in bus:
        cb = cur["coll_bu"]
        ar["collections_bu"].append([b, cb.get((b, "month_prev"), 0.0), cb.get((b, "month"), 0.0), cb.get((b, "prior"), 0.0), cb.get((b, "ytd_prev"), 0.0), cb.get((b, "ytd"), 0.0), cur["coll_cnt"].get(b, 0), cur["deposited_bu"].get(b, 0.0), cb.get((b, "month_prev_nonocp"), 0.0), cb.get((b, "month_nonocp"), 0.0), cb.get((b, "prior_nonocp"), 0.0), cb.get((b, "ytd_nonocp"), 0.0)])
    ar["ap_bu"] = [[b, ap_prev_bu.get(b, {}).get("downpayment", 0.0), ap_cur_bu[b]["downpayment"], ap_prev_bu.get(b, {}).get("parked", 0.0), ap_cur_bu[b]["parked"],
                    ap_prev_bu.get(b, {}).get("inv4411", 0.0), ap_cur_bu[b]["inv4411"], ap_prev_bu.get(b, {}).get("lc", 0.0), ap_cur_bu[b]["lc"]] for b in bus if b in ap_cur_bu or b in ap_prev_bu]
    # per-BU detail (current / prev when a raw prev file or prev inputs with detail exist / month start)
    detail = []
    Pd = P.get("detail_agg"); Pdet = P.get("detail")
    def pv(section, b, item, key):
        if Pd: return {"unbilled_status": Pd["status_bu"].get((b, key), 0.0), "unpaid_group": Pd["unpaid_group"].get((b, key), 0.0), "unpaid_phase": Pd["unpaid_phase"].get((b, key), 0.0),
                       "aging_unbilled": Pd["aging_unbilled"].get((b, key), 0.0), "aging_unpaid": Pd["aging_unpaid"].get((b, key), 0.0), "owner": Pd["owner"].get((b, key), 0.0), "advance_status": Pd["advance_status"].get((b, key), 0.0)}[section]
        if Pdet is not None: return Pdet.get((b, section, item))
        if section == "owner": return P["owner"].get((b, key))
        return None
    def mv(section, b, key):
        if not MS: return None
        return {"unbilled_status": MS["status_bu"], "unpaid_group": MS["unpaid_group"], "unpaid_phase": MS["unpaid_phase"], "aging_unbilled": MS["aging_unbilled"], "aging_unpaid": MS["aging_unpaid"], "owner": MS["owner"], "advance_status": MS["advance_status"]}[section].get((b, key), 0.0)
    for b in bus:
        for cat in ("At risk", "Under JESA", "Under OCP", "SS postponed"):
            detail.append([b, "owner", cat, mv("owner", b, cat), pv("owner", b, cat, cat), cur["owner"].get((b, cat), 0.0)])
        for c in sorted({k[1] for k in cur["status_bu"] if k[0] == b} | ({k[1] for k in MS["status_bu"] if k[0] == b} if MS else set()), key=lambda c: -cur["status_bu"].get((b, c), 0)):
            detail.append([b, "unbilled_status", status_label(c), mv("unbilled_status", b, c), pv("unbilled_status", b, status_label(c), c), cur["status_bu"].get((b, c), 0.0)])
        for gname in ("Current", "Past due", "Rejected", "Other"):
            detail.append([b, "unpaid_group", gname, mv("unpaid_group", b, gname), pv("unpaid_group", b, gname, gname), cur["unpaid_group"].get((b, gname), 0.0)])
        for ph in ("Rejected", "Treasury", "OCP subs", "Ordonnancement", "Missing reception", "Non-OCP / OCP sub (hors)", "Being processed", "Other"):
            detail.append([b, "unpaid_phase", ph, mv("unpaid_phase", b, ph), pv("unpaid_phase", b, ph, ph), cur["unpaid_phase"].get((b, ph), 0.0)])
        for bk in BUCKETS:
            detail.append([b, "aging_unbilled", bk, mv("aging_unbilled", b, bk), pv("aging_unbilled", b, bk, bk), cur["aging_unbilled"].get((b, bk), 0.0)])
            detail.append([b, "aging_unpaid", bk, mv("aging_unpaid", b, bk), pv("aging_unpaid", b, bk, bk), cur["aging_unpaid"].get((b, bk), 0.0)])
        for lab in sorted({k[1] for k in cur["advance_status"] if k[0] == b}):
            detail.append([b, "advance_status", lab, mv("advance_status", b, lab), pv("advance_status", b, lab, lab), cur["advance_status"].get((b, lab), 0.0)])
    projects = [[b, t, w, n, v, st] for (b, t), lst in cur["top_projects"].items() for (w, n, v, st) in lst]

    meta = {"as_of_date": as_of, "week_label": a.week_label or f"Week {as_of.isocalendar()[1]}", "ar_current_date": ar_date, "ar_prev_date": prev_date,
            "ar_month_start_date": ms_date, "cash_data_as_of": cash["data_as_of"].date() if cash["data_as_of"] else None,
            "collections_to": a.collections_to or "19/10", "payments_to": a.payments_to or "15/10", "lc_date": a.lc_date or "19/10", "ap_label": a.ap_label or "Oct W3",
            "cash_basis": cash["basis"], "units": "MMAD (millions of MAD)", "ur_prev_source": (ur or {}).get("prev_source", ""), "collections_basis": "AR paid lines by payment date, amount invoiced (same basis as 'Invoices paid' in the BU meeting report)",
            "deposited_month": cur["deposited_total"]}

    ar["jesa_unbilled_tot"] = (sum(cur["jesa_unbilled"].values()), sum(v for (b, l), v in cur["jesa_unbilled"].items() if l == "labor"), sum(v for (b, l), v in cur["jesa_unbilled"].items() if l == "non_labor"))
    ar["nonocp_tot"] = sum(cur["nonocp"].values())
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    write_sheet(wb, "META", ["key", "value"], [[k, (v.isoformat() if isinstance(v, dt.date) else v)] for k, v in meta.items()], note="Dates and labels used on the slides. Edit freely.", widths=[24, 60])
    write_sheet(wb, "KEY_MESSAGES", ["n", "title", "body"], [[i + 1, t, b] for i, (t, b) in enumerate(default_messages(ar, cash, lc, ap, meta))], note="Drafted from the figures in deck order; rewrite the wording as needed (max 6 messages).", widths=[4, 70, 120])
    write_sheet(wb, "AR_KPI", ["metric", "month_start", "prev_week", "current"], ar["kpi"], note="Key figures (MMAD): AR from the raw AR extract (dashboard rules), unrecognized from the UR consolidation, AP from the supplier invoice file, LC from the bank LC situation.", widths=[44])
    ls = ar["labor_split"]
    write_sheet(wb, "AR_SPLITS", ["metric", "value"], [["Total AR labor", ls["total_labor"]], ["Total AR non-labor", ls["total_nonlabor"]], ["Unbilled labor", ls["unbilled_labor"]], ["Unbilled non-labor", ls["unbilled_nonlabor"]], ["Unpaid labor", ls["unpaid_labor"]], ["Unpaid non-labor", ls["unpaid_nonlabor"]]] + [[f"AR {k}", v] for k, v in ar["ocp_split"].items()], note="Current week only (labor = PSR/PROC column; OCP = Client 1).", widths=[30])
    write_sheet(wb, "AR_BU", ["bu", "component", "month_start", "prev_week", "current"], ar["ar_bu"] + [[b, "Unrecognized", (ur_ms.get(b) if ur_ms else None), (ur["prev"].get(b) if ur else None), v] for b, v in ar["unrecognized_bu"].items()], note="AR by business unit and component (MMAD).", widths=[28, 16])
    write_sheet(wb, "AR_UNBILLED_OWNER", ["bu", "category", "month_start", "prev_week", "current"], ar["unbilled_owner"], note="Unbilled split by ownership: At risk / Under JESA / Under OCP / Service sheets postponed.", widths=[28, 16])
    write_sheet(wb, "AR_UNBILLED_STATUS", ["group", "status", "month_start", "prev_week", "current"], ar["unbilled_status"], note="Unbilled by coded status (all BUs).", widths=[10, 50])
    write_sheet(wb, "AR_UNPAID_STATUS", ["status", "month_start", "prev_week", "current"], ar["unpaid_status"], note="Unpaid by OCP processing phase (from the 'OCP Comments' column).", widths=[28])
    write_sheet(wb, "AR_AGING", ["year", "month_start", "prev_week", "current"], ar["aging"], note="Total AR by year of origin.", widths=[24])
    write_sheet(wb, "AR_BU_DETAIL", ["bu", "section", "item", "month_start", "prev_week", "current"], detail, note="Per-BU detail for the zoom slides: owner, unbilled_status, unpaid_group, unpaid_phase, aging_unbilled, aging_unpaid, advance_status (MMAD). prev_week filled when last week's raw AR file or inputs workbook with detail is available.", widths=[28, 18, 50])
    write_sheet(wb, "AR_BU_PROJECTS", ["bu", "type", "wbs", "project", "amount", "status"], projects, note="Top 8 unbilled and unpaid projects per BU (MMAD, current week).", widths=[28, 10, 14, 50, 10, 44])
    write_sheet(wb, "UNRECOGNIZED", ["bu", "eq_pending_approval", "eq_signed_by_client", "eq_pending_submission", "eq_under_cn", "eq_at_risk", "sub_pending_approval", "sub_signed_by_client", "sub_pending_submission", "sub_under_cn", "sub_at_risk", "total_suppliers", "lines_to_adjust", "grand_total", "prev_week_total", "month_start_total"],
                [r[:14] + [(ur["prev"].get(r[0]) if ur else None), (ur_ms.get(r[0]) if ur_ms else None)] for r in ar["unrecognized_table"]], note=f"Unrecognized suppliers by status, MMAD incl. markup; prev_week_total = {(ur or {}).get('prev_source', 'n/a')}.", widths=[28] + [14] * 14)
    if ur:
        write_sheet(wb, "UR_ATTACHMENT_STATUS", ["bu", "eq_under_jesa_approval", "eq_under_jesa_preparation", "eq_under_client_approval", "eq_signed_by_client", "eq_accrual_posted", "sub_under_jesa_approval", "sub_under_jesa_preparation", "sub_under_client_approval", "sub_signed_by_client", "sub_accrual_posted", "total"], ur["attachment"], note="Status of JESA attachments to client (UR consolidation, MMAD incl. markup).", widths=[28] + [14] * 11)
        write_sheet(wb, "UR_PO_SUMMARY", ["bu", "contract_value", "recognized", "unrecognized", "remaining"], ur["po"], note="Supplier PO consolidation by BU (MMAD).", widths=[28, 16, 16, 16, 16])
    Pn = P.get("nonocp", {}); Pj = P.get("jesa_unbilled", {})
    write_sheet(wb, "NONOCP_BU", ["bu", "component", "month_start", "prev_week", "current"],
                [[b, c, (MS["nonocp"].get((b, c), 0.0) if MS else None), Pn.get((b, c)), cur["nonocp"].get((b, c), 0.0)] for b in bus for c in ("Unbilled", "Unpaid")],
                note="Non-OCP clients (Client 1 = NON OCP): unbilled and unpaid by BU (MMAD).", widths=[28, 12])
    write_sheet(wb, "NONOCP_PROJECTS", ["bu", "client", "wbs", "project", "unbilled", "unpaid", "status"], cur["nonocp_projects"][:40],
                note="Non-OCP projects driving unbilled and unpaid (top 40 by total, MMAD, current week).", widths=[26, 30, 14, 44, 10, 10, 40])
    write_sheet(wb, "JESA_UNBILLED_BU", ["bu", "split", "month_start", "prev_week", "current"],
                [[b, sp, (MS["jesa_unbilled"].get((b, sp), 0.0) if MS else None), Pj.get((b, sp)), cur["jesa_unbilled"].get((b, sp), 0.0)] for b in bus for sp in ("labor", "non_labor")],
                note="Unbilled where the action sits with JESA (J-* statuses except SS postponed), by BU and PSR (labor) / Procurement (non-labor), MMAD.", widths=[28, 12])
    from ar_engine import status_label as _sl
    jcodes = sorted({c for (c, l) in cur["jesa_unbilled_status"]}, key=lambda c: -(cur["jesa_unbilled_status"].get((c, "labor"), 0) + cur["jesa_unbilled_status"].get((c, "non_labor"), 0)))
    write_sheet(wb, "JESA_UNBILLED_STATUS", ["status", "labor", "non_labor"], [[_sl(c), cur["jesa_unbilled_status"].get((c, "labor"), 0.0), cur["jesa_unbilled_status"].get((c, "non_labor"), 0.0)] for c in jcodes],
                note="Unbilled under JESA responsibility by status and PSR / Procurement (MMAD, current week).", widths=[50, 12, 12])
    write_sheet(wb, "DSO", ["bu", "dso_current", "dso_prior_month", "total_ar", "revenue_12m"], [[b, v.get("current"), v.get("prior"), v.get("ar"), v.get("rev12")] for b, v in dso.items()],
                note="DSO (days) = total AR of the BU / 12-month trailing revenue x 365 (DSO calc workbook); 'JESA Group' row = group DSO; prior month from the Trend sheet.", widths=[28, 12, 14, 12, 12])
    if dso_trend:
        write_sheet(wb, "DSO_TREND", ["bu"] + dso_trend["trend_months"] + ["current"], [[b] + (dso_trend["trend"].get(b) or [None] * len(dso_trend["trend_months"])) + [dso[b]["current"]] for b in dso if b != "JESA Group"],
                    note=f"DSO history by BU (days) from the DSO calc workbook 'Trend' sheet; 12-month revenue window {dso_trend['window'][0]} to {dso_trend['window'][-1]}.", widths=[28] + [9] * (len(dso_trend["trend_months"]) + 1))
    write_sheet(wb, "COLLECTIONS_BU", ["bu", "month_prev_week", "month_current", "prior_month", "ytd_prev_week", "ytd_current", "invoices_ytd", "deposited_month", "nonocp_month_prev_week", "nonocp_month_current", "nonocp_prior_month", "nonocp_ytd"], sorted(ar["collections_bu"], key=lambda r: -r[5]),
                note=f"Payments received by BU (amount invoiced, by payment date): current month to {ar_date.isoformat()} vs to {prev_date.isoformat()}, prior full month, YTD; deposited_month = invoices received (reception date) in the current month, HT; nonocp_* = same measures for non-OCP clients (Client 1).", widths=[28])
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    write_sheet(wb, "COLLECTIONS_MONTHLY", ["month", "labor", "non_labor"], [[m, cur["coll_month"].get((i + 1, "labor"), 0.0), cur["coll_month"].get((i + 1, "non_labor"), 0.0)] for i, m in enumerate(months)],
                note="Monthly payments received (amount invoiced, by payment date), labor (PSR) vs non-labor (PROC).", widths=[10])
    write_sheet(wb, "CASH_MONTHLY", ["line", "jul", "aug", "sep", "oct_actual", "oct_ongoing", "cumulative"], cash["cash_monthly"], note="Cash flow workbook 'EXEC Summary' (MMAD): cash in positive, cash out negative.", widths=[48])
    write_sheet(wb, "CASH_PSR_BU", ["bu", "jul", "aug", "sep", "oct_actual", "oct_ongoing", "cumulative"], cash.get("psr_bu", []), note="PSR (labor) collections by BU per OCP payment report (EXEC Summary section 3), MMAD.", widths=[28])
    write_sheet(wb, "CASH_BU", ["bu", "cash_in", "cash_out_suppliers", "cash_out_lc", "cash_out_customs", "net_actuals", "oct_approved", "net_after_approved", "oct_pending", "net_all_ongoing"], cash["cash_bu"], note="Procurement net balance by BU, Jul–Oct cumulative (EXEC Summary section 5).", widths=[28])
    bu_oct = []
    for b in BU_ORDER:
        if b in cash["bu_cash_in"] or b in cash["bu_cash_out"]:
            ci = cash["bu_cash_in"].get(b, [0] * 6); co = cash["bu_cash_out"].get(b, [0] * 6); lcv = cash["bu_lc"].get(b, [0] * 6); cu = cash["bu_customs"].get(b, [0] * 6)
            bu_oct.append([b, ci[3], co[3], lcv[3], cu[3], co[4], lcv[4]])
    write_sheet(wb, "CASH_BU_OCT", ["bu", "oct_cash_in", "oct_paid_suppliers", "oct_paid_lc", "oct_paid_customs", "oct_ongoing_suppliers", "oct_ongoing_lc"], bu_oct, note="Current month only: collections and payments to date by BU, plus open commitments.", widths=[28])
    cm = {r[0]: r for r in cash["cash_monthly"]}
    tot_appr = sum(r[6] for r in cash["cash_bu"]); tot_pend = sum(r[8] for r in cash["cash_bu"])
    if not cash.get("lc_oct_maturity_total") and lc.get("months"):
        # 'LC Ongoing' sheet absent: take the current-month maturities from the bank LC situation instead
        lc_oct_m = sum(l["months"].get(lc["months"][0], 0) for l in lc["lines"])
        cash["lc_oct_maturity_total"] = min(lc_oct_m, abs(tot_appr))
        cash["lc_oct_maturity_bu"] = {}
        print(f"INFO LC {lc['months'][0]} maturities taken from the LC situation file: {cash['lc_oct_maturity_total']:,.1f} MMAD")
    ongoing = [["Paid in October to date (suppliers + LC + customs)", None, abs(cm["Total cash OUT (procurement)"][4])], ["Approved – October LC maturities", None, cash["lc_oct_maturity_total"]],
               ["Approved – supplier invoices", None, abs(tot_appr) - cash["lc_oct_maturity_total"]], ["Pending payment approval", None, abs(tot_pend)]]
    ongoing += [[f"Tracking log – {r[0]}", r[1], r[2]] for r in cash["approved_status"]]
    ongoing += [["Priority list – total approved", None, cash["priority_total"]], ["Priority list – carried in October ongoing", None, cash["priority_carried"]]]
    ongoing += [[f"Priority list – {k}", None, v] for k, v in cash["priority_by_class"].items()]
    write_sheet(wb, "ONGOING", ["bucket", "lines", "amount"], ongoing, note="October ongoing payments position (MMAD, positive = to pay).", widths=[60, 8])
    write_sheet(wb, "AP_BU", ["bu", "downpayment_prev", "downpayment_cur", "parked_prev", "parked_cur", "inv4411_prev", "inv4411_cur", "lc_prev", "lc_cur"], ar["ap_bu"],
                note="Supplier payables by BU (supplier invoice file) and opened LC by BU (bank LC situation), MMAD, previous week vs current.", widths=[28])
    write_sheet(wb, "AP_BUCKETS", ["origin", "bucket", "amount"], [[k.split("|")[0], k.split("|")[1], v] for k, v in sorted(ap["buckets"].items())], note=f"Supplier invoice file ({meta['ap_label']}): open payables by origin and due-date bucket vs {as_of.isoformat()} (MMAD).", widths=[20, 16])
    write_sheet(wb, "AP_BU_ORIGIN", ["bu", "origin", "amount", "lines"], [[k.split("|")[0], k.split("|")[1], v, ap["count"][k]] for k, v in sorted(ap["by_bu_origin"].items())], note="Supplier invoice file: open payables by BU and origin.", widths=[28, 20])
    write_sheet(wb, "AP_TIERS", ["tier", "amount"], sorted([[k, v] for k, v in ap["tiers"].items()], key=lambda r: -r[1]), note="Supplier invoice file: open payables by supplier tier.", widths=[24])
    write_sheet(wb, "AP_TOP_SUPPLIERS", ["supplier", "amount", "bu"], ap["top_suppliers"], note="Top 15 suppliers by open payables.", widths=[36, 12, 28])
    write_sheet(wb, "LC_LINES", ["project", "bu", "supplier", "lc_ref", "expiry", "ccy", "face_value_fx", "encours_fx", "encours_mad", "opened", "action", "part_2026", "part_2027"] + lc["months"],
                [[l["project"], l["bu"], l["supplier"], l["ref"], l["expiry"], l["ccy"], l["face_fx"], l["encours_fx"], l["encours_mad"], l["opened"], l["action"], l["part_2026"], l["part_2027"]] + [l["months"][m] for m in lc["months"]] for l in lc["lines"]],
                note=f"LC situation {meta['lc_date']}: one line per LC (MMAD).", widths=[14, 24, 30, 18, 12, 6, 14, 14, 12, 8, 36, 10, 10])
    wb.save(a.out)
    json.dump({"meta": {k: (v.isoformat() if isinstance(v, dt.date) else v) for k, v in meta.items()}}, open(a.out.replace(".xlsx", "_meta.json"), "w"), indent=1, default=str)
    print("written", a.out)

if __name__ == "__main__":
    main()
