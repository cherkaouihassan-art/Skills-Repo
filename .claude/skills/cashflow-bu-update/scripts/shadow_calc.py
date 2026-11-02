#!/usr/bin/env python3
"""Recompute 'REPORT - Cashflow by BU' from the source tabs (same SUMIFS logic as the workbook) and compare
with the cached values of the report. Works without Excel/LibreOffice.
Usage: shadow_calc.py <workbook.xlsx> [--json out.json] [--tol 1] [--quiet]
Exit code 0 when every compared cell is within tolerance, 2 otherwise."""
import sys, os, json, datetime as dt
from collections import defaultdict
from openpyxl import load_workbook
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfcommon import SHEETS, BUS, TIERS, LC_PROJECT, MONTHS, REPORT_COLS, BU_CODES, to_date, month_start, fmt

path = sys.argv[1]
tol = float(sys.argv[sys.argv.index('--tol') + 1]) if '--tol' in sys.argv else 1.0
quiet = '--quiet' in sys.argv
wb = load_workbook(path, read_only=True, data_only=True)
OCT = dt.date(2026, 10, 1)

ERR = ('#NAME?', '#REF!', '#VALUE!', '#N/A', '#DIV/0!', '#NUM!', '#NULL!')
def clean(v): return None if (isinstance(v, str) and v.strip() in ERR) else v
def rows(sheet, maxc):
    for r in wb[sheet].iter_rows(min_row=2, max_col=maxc, values_only=True):
        yield tuple(clean(v) for v in r)

def col(letter):
    from openpyxl.utils import column_index_from_string as ci
    return ci(letter) - 1

S = defaultdict(float)   # (key tuple) -> amount

# ---- Cash IN (OCP collections) : Z amount, AD pay date, AO BU, AP Labor flag
for r in rows(SHEETS['cashin'], 42):
    d = to_date(r[col('AD')]); 
    if d is None: continue
    m = month_start(d); amt = r[col('Z')]
    if not isinstance(amt, (int, float)): continue
    bu, flag = r[col('AO')], r[col('AP')]
    if flag in ('Labor', 'Non-Labor'): S[('in', flag, 'ALL', m)] += amt
    if bu not in (None, '', 'Unmapped'):
        S[('in', 'mapped', 'ALL', m)] += amt
        S[('in', 'mapped', bu, m)] += amt
        if flag in ('Labor', 'Non-Labor'): S[('in', flag, bu, m)] += amt

# ---- SAP payments Jul / Aug (L amount, A BU, B tier, O OH)
OHN0 = set()
for row in wb[SHEETS['oh']].iter_rows(min_row=2, values_only=True):
    if len(row) > 4 and row[4]: OHN0.add(__import__('cfcommon').norm_supplier(row[4]))
for key, m in (('pay_jul', MONTHS[0]), ('pay_aug', MONTHS[1])):
    for r in rows(SHEETS[key], 16):
        bu, tier, amt, oh = r[0], r[1], r[11], r[14]
        if oh is None: oh = 'OH' if __import__('cfcommon').norm_supplier(r[5]) in OHN0 else ''
        if not isinstance(amt, (int, float)) or bu in (None, ''): continue
        t = str(tier)
        S[('out', t, 'ALL', m)] += amt; S[('out', t, bu, m)] += amt
        if oh == 'OH': S[('oh', 'ALL', m)] += amt; S[('oh', bu, m)] += amt
# ---- September (C amount, AK BU, AM tier, AP OH); AM/AK/AP are formulas -> re-derive when not calculated
from cfcommon import norm_supplier as _ns
TM = {}
for row in wb[SHEETS['tiering']].iter_rows(min_row=5, values_only=True):
    if row[0]: TM[_ns(row[0])] = str(row[2] or '').strip()
OHN = set()
for row in wb[SHEETS['oh']].iter_rows(min_row=2, values_only=True):
    if len(row) > 4 and row[4]: OHN.add(_ns(row[4]))
NAT = {'1 - National': '1', '2 - National': '2', '3 - National': '3', '1 - International': 'Intl 1', '2 - International': 'Intl 2', '3 - International': 'Intl 3'}
for r in rows(SHEETS['pay_sep'], 42):
    amt, bu, tier, oh = r[2], r[col('AK')], r[col('AM')], r[col('AP')]
    if bu is None:
        ai = str(r[col('AI')] or '').strip(); bu = '(unmapped)' if ai == '' else ('JESA Technologies' if ai.upper() == 'JESA TECHNOLOGIES' else ai)
    if tier is None:
        a = str(r[0] or '').strip()
        if a in NAT: tier = NAT[a]
        elif a == 'International': t0 = TM.get(_ns(r[5]), ''); tier = 'Intl ' + (t0 if t0 not in ('', '0') else '1')
        else: tier = a
    if oh is None: oh = 'OH' if _ns(r[5]) in OHN else ''
    if not isinstance(amt, (int, float)) or bu in (None, ''): continue
    t = str(tier); m = MONTHS[2]
    if bu != '(unmapped)': S[('out', t, 'ALL', m)] += amt
    S[('out', t, bu, m)] += amt
    if oh == 'OH':
        if bu != '(unmapped)': S[('oh', 'ALL', m)] += amt
        S[('oh', bu, m)] += amt
# ---- October (C amount positive -> negate, AH BU, AG tier, AJ OH, AF include, G name)
for r in rows(SHEETS['pay_oct'], 37):
    if r[col('AF')] != 'Yes': continue
    amt, bu, tier, oh, name = -(r[2] or 0), r[col('AH')], str(r[col('AG')]), r[col('AJ')], str(r[6] or '')
    m = OCT
    if bu != '(unmapped)': S[('out', tier, 'ALL', m)] += amt
    else: S[('out', 'UNMAPPED', 'ALL', m)] += amt
    S[('out', tier, bu, m)] += amt
    if oh == 'OH': S[('oh', 'ALL', m)] += amt; S[('oh', bu, m)] += amt
    if bu == 'AM' and name.upper().startswith('PISTELLI'): S[('pistelli_paid',)] += -amt   # report adds back positive

# ---- LC actuals (E date, G debit, S project, X vendor)
for r in rows(SHEETS['lc'], 37):
    d = to_date(r[4]); amt = r[6]
    if d is None or not isinstance(amt, (int, float)): continue
    m = month_start(d); proj = r[col('S')]
    S[('lc', proj, m)] += -amt
    if proj == 'AMT' and str(r[col('X')] or '').upper().startswith('PELZ') and m == OCT: S[('pelz_lc_oct',)] += amt
# ---- LC ongoing (P BU, Q signed) rows 5:32
LC_TAB = sys.argv[sys.argv.index('--lc-tab') + 1] if '--lc-tab' in sys.argv else SHEETS['lc_ongoing']
lo = wb[LC_TAB]
for rr in range(5, lo.max_row + 1):
    if str(lo.cell(rr, 1).value or '').upper().startswith('TOTAL') or lo.cell(rr, 3).value is None: 
        if lo.cell(rr, 16).value is None: continue
    bu, q, R, Lm = (clean(lo.cell(rr, c).value) for c in (16, 17, 18, 12))
    if bu is None and lo.cell(rr, 3).value:   # P is an XLOOKUP (not evaluated by LibreOffice): derive from the project code
        code = str(lo.cell(rr, 3).value).strip().upper(); bu = BU_CODES.get(code) or BU_CODES.get(code.split('-')[0])
    if bu in ('Business Unit', None) or str(lo.cell(rr, 1).value or '').upper().startswith('TOTAL'): continue
    if q is None and isinstance(Lm, (int, float)): q = 0 if str(R) == 'Yes' else -Lm   # formula not yet calculated
    if bu and isinstance(q, (int, float)): S[('lc_ongoing', bu)] += q
# ---- Customs DD (E amount, I BU, J month); rows added by openpyxl have no cached I/J -> derive from H/B
for r in rows(SHEETS['customs'], 13):
    amt, bu, m = r[4], r[8], to_date(r[9])
    if bu is None and r[7]: bu = BU_CODES.get(str(r[7]).strip().upper(), str(r[7]).strip())
    if m is None and r[1]: m = month_start(r[1])
    if isinstance(amt, (int, float)) and bu and m: S[('customs', bu, m)] += -amt
# ---- Ongoing position: E amount, A BU, B supplier, G tier, N status, P OH, AA approved
OH_NAMES = set()
from cfcommon import norm_supplier
for row in wb[SHEETS['oh']].iter_rows(min_row=2, values_only=True):
    for v in (row[3] if len(row) > 3 else None, row[4] if len(row) > 4 else None):
        if v: OH_NAMES.add(norm_supplier(v))
for r in rows(SHEETS['ongoing'], 34):
    bu, sup, amt, tier, status, oh, appr = r[0], str(r[1] or ''), r[4], str(r[6]), r[13], r[15], r[26]
    if status is None:   # row added by the loader: status/OH/approved are formulas without cached values yet
        status = str(r[16] or 'Open'); oh = 'OH' if norm_supplier(sup) in OH_NAMES else ''
    if appr is None: appr = 'Yes' if str(r[20]) == 'Yes' else 'No'
    status = str(status or '')
    if bu is None or not isinstance(amt, (int, float)) or not status.startswith('Open'): continue
    flag = 'F' if appr == 'Yes' else 'H'
    S[('og', flag, tier, 'ALL')] += amt; S[('og', flag, tier, bu)] += amt
    if oh == 'OH': S[('og_oh', flag, 'ALL')] += amt; S[('og_oh', flag, bu)] += amt
    if bu == 'AM' and sup.upper().startswith('PISTELLI'): S[('og_pistelli', flag)] += amt
# ---- standing inputs
pc = wb[SHEETS['proc_comp']]
DIV = {pc.cell(r, 1).value: (clean(pc.cell(r, 3).value) or 0) for r in range(5, 15)}
DIV_TOTAL = clean(pc.cell(15, 3).value) or sum(v for v in DIV.values() if isinstance(v, (int, float)))

def tier_out(tier, bu, m):
    if tier == 'Unclassified / Not on OCP':
        v = S[('out', 'Not on OCP Tier', bu, m)]
        if bu == 'ALL' and m == OCT: v += S[('out', 'UNMAPPED', 'ALL', m)]
        return v
    return S[('out', tier, bu, m)]

def tier_og(flag, tier, bu):
    if tier == 'Unclassified / Not on OCP':
        return S[('og', flag, 'Unclassified', bu)] + S[('og', flag, 'Not on OCP Tier', bu)]
    return S[('og', flag, tier, bu)]

# ---- assemble expected values per block (label -> {col: value})
ALL_TIERS = TIERS + ['Unclassified / Not on OCP']
def block(bu):
    E = defaultdict(dict); key = bu if bu != 'JESA' else 'ALL'
    for m in MONTHS:
        c = REPORT_COLS[m]
        psr = S[('in', 'Labor', key, m)]; proc = S[('in', 'Non-Labor', key, m)]
        uncl = S[('in', 'mapped', key, m)] - psr - proc
        div = (DIV_TOTAL if bu == 'JESA' else DIV.get(bu, 0)) if m == MONTHS[2] else 0
        E['PSR (Labor)'][c] = psr; E['Procurement (Non-Labor)'][c] = proc
        E['Unclassified'][c] = uncl; E['Dividends'][c] = div; E['Total Cash IN'][c] = proc + div + uncl
        for t in ALL_TIERS: E[t][c] = tier_out(t, key, m)
        if bu == 'JESA': lc = sum(S[('lc', LC_PROJECT[b], m)] for b in LC_PROJECT); cu = sum(S[('customs', b, m)] for b in BUS)
        else: lc = S[('lc', LC_PROJECT[bu], m)] if bu in LC_PROJECT else 0; cu = S[('customs', bu, m)]
        E['LC payments'][c] = lc; E['Customs'][c] = cu
        E['of which Overhead (OH)'][c] = S[('oh', key, m)]
        E['Less OH AM'][c] = -S[('oh', 'AM', m)]; E['Less OH Water'][c] = -S[('oh', 'Water & Environement', m)]
        E['Pistelli'][c] = (S[('pistelli_paid',)] + S[('pelz_lc_oct',)]) if m == OCT else 0
    for flag in ('F', 'H'):
        for t in ALL_TIERS: E[t][flag] = tier_og(flag, t, key)
        E['of which Overhead (OH)'][flag] = S[('og_oh', flag, key)]
        E['Less OH AM'][flag] = -S[('og_oh', flag, 'AM')]; E['Less OH Water'][flag] = -S[('og_oh', flag, 'Water & Environement')]
        E['Pistelli'][flag] = -S[('og_pistelli', flag)]
        E['Total Cash IN'][flag] = 0
    E['LC payments']['F'] = sum(S[('lc_ongoing', b)] for b in BUS) if bu == 'JESA' else S[('lc_ongoing', bu)]
    E['LC payments']['H'] = 0; E['Customs']['F'] = 0; E['Customs']['H'] = 0
    # totals
    for c in ['B', 'C', 'D', 'E', 'F', 'H']:
        out = sum(E[t][c] for t in ALL_TIERS) + E['LC payments'][c] + E['Customs'][c]
        if bu in ('JESA', 'AM'): out += E['Less OH AM'][c] + E['Pistelli'][c]
        if bu in ('JESA', 'Water & Environement'): out += E['Less OH Water'][c]
        E['Total Cash OUT (incl. LC & customs)'][c] = out
        E['NET BALANCE'][c] = E['Total Cash IN'][c] + out
    return E

LABEL_MAP = {'PSR (Labor)': 'PSR (Labor)', 'Procurement (Non-Labor)': 'Procurement (Non-Labor)',
             'Cash IN — PROC Compensable (standing)': 'Dividends', 'Unclassified collections': 'Unclassified',
             'Total Cash IN': 'Total Cash IN', 'LC payments': 'LC payments', 'Customs': 'Customs',
             'Total Cash OUT (incl. LC & customs)': 'Total Cash OUT (incl. LC & customs)', 'NET BALANCE': 'NET BALANCE',
             'of which Overhead (OH)': 'of which Overhead (OH)'}
for t in ALL_TIERS: LABEL_MAP[t] = t

ws = wb[SHEETS['report']]
heads = {'JESA': 7, **{b: r for b, r in __import__('cfcommon').BLOCK_ROWS.items()}}
order = sorted(heads.items(), key=lambda x: x[1]); ends = {k: (order[i + 1][1] if i + 1 < len(order) else 336) for i, (k, _) in enumerate(order)}
report = {}; diffs = []; ncomp = 0
for bu, h in order:
    E = block(bu); report[bu] = {}
    for r in range(h, ends[bu]):
        lab = ws.cell(r, 1).value
        if lab is None: continue
        lab = str(lab).strip()
        k = LABEL_MAP.get(lab) or LABEL_MAP.get(lab.split(' (')[0] if lab.startswith('Unclassified') else lab)
        if lab.startswith('Less Overhead') or lab.startswith('Less AM Overhead'): k = 'Less OH AM' if bu in ('JESA', 'AM') else 'Less OH Water'
        if lab.startswith('Less Water Overhead'): k = 'Less OH Water'
        if lab.startswith('Less PISTELLI') or lab.startswith('Less AM PISTELLI'): k = 'Pistelli'
        if lab.startswith('LC payments'): k = 'LC payments'
        if lab.startswith('Customs'): k = 'Customs'
        if lab.startswith('Unclassified collections'): k = 'Unclassified'
        if k is None or k not in E: continue
        for c in ['B', 'C', 'D', 'E', 'F', 'H']:
            if c not in E[k]: continue
            if k == 'of which Overhead (OH)':   # memo line: legacy formulas differ per block, compare only where consistent
                if bu == 'JESA' and c == 'H': continue
                if bu == 'JESA' and c == 'F': E[k]['F'] = S[('og_oh', 'F', 'ALL')] + S[('og_oh', 'H', 'ALL')]
                if bu not in ('JESA', 'AM', 'Water & Environement') and c in ('E', 'F', 'H'): continue
            cached = clean(ws[f'{c}{r}'].value); exp = E[k][c]
            cached = float(cached) if isinstance(cached, (int, float)) else 0.0
            report[bu][f'{k}|{c}'] = {'expected': exp, 'cached': cached, 'cell': f'{c}{r}'}
            ncomp += 1
            if abs(exp - cached) > tol: diffs.append((bu, k, f'{c}{r}', exp, cached, exp - cached))

# ---- headline from shadow values
J = block('JESA'); sw = 1 if str(wb[SHEETS['exec']]['F6'].value).upper() == 'TRUE' else 0
head = {}
head['collections_ocp'] = sum(J['Procurement (Non-Labor)'][c] + J['Unclassified'][c] for c in 'BCDE')
head['dividends'] = DIV_TOTAL
head['cash_out_jul_sep'] = sum(J['Total Cash OUT (incl. LC & customs)'][c] for c in 'BCD')
head['cash_out_oct_actual'] = J['Total Cash OUT (incl. LC & customs)']['E']
head['approved_pmts'] = J['Total Cash OUT (incl. LC & customs)']['F'] * sw
head['net_after_approved'] = sum(head[k] for k in ['collections_ocp', 'dividends', 'cash_out_jul_sep', 'cash_out_oct_actual', 'approved_pmts'])
head['pending_pmts'] = J['Total Cash OUT (incl. LC & customs)']['H'] * sw
head['net_all_ongoing'] = head['net_after_approved'] + head['pending_pmts']
ex = wb[SHEETS['exec']]
head_cached = {k: clean(ex.cell(13, 3 + i).value) for i, k in enumerate(head)}
if '--baseline' in sys.argv:   # compare with a previous state/shadow json instead of the (possibly stale) cached cover
    b = json.load(open(sys.argv[sys.argv.index('--baseline') + 1]))
    head_cached = b.get('headline_mad') or b.get('headline_shadow') or head_cached

out = {'compared_cells': ncomp, 'differences': [dict(zip(['bu', 'line', 'cell', 'expected', 'cached', 'diff'], d)) for d in diffs],
       'headline_shadow': head, 'headline_cached': head_cached,
       'headline_diff': {k: (head[k] - (head_cached[k] or 0)) for k in head}}
if '--json' in sys.argv: json.dump(out, open(sys.argv[sys.argv.index('--json') + 1], 'w'), indent=1, default=str)
if not quiet:
    print(f"Compared {ncomp} report cells against the shadow calculation; {len(diffs)} outside ±{tol} MAD" + (" (cached values are stale after a load: each difference should be a loaded line)" if '--baseline' in sys.argv else ''))
    for d in diffs[:60]: print(f"  {d[0]:26s} {d[1]:36s} {d[2]:6s} expected {fmt(d[3]):>18s} cached {fmt(d[4]):>18s} diff {fmt(d[5]):>14s}")
    print("\nHEADLINE  shadow vs " + ('baseline' if '--baseline' in sys.argv else 'cached'))
    for k in head: print(f"  {k:22s} {fmt(head[k]):>20s} {fmt(head_cached[k] or 0):>20s} {fmt(head[k]-(head_cached[k] or 0)):>14s}")
sys.exit(0 if not diffs and all(abs(v) <= tol for v in out['headline_diff'].values()) else 2)
