#!/usr/bin/env python3
"""Load an intake pack (CF_INPUT_PACK_<date>.xlsx) into the CASH FLOW ANALYSIS - BUs master.

Usage:
  load_pack.py <master.xlsx> <pack.xlsx> --out <new_master.xlsx> --as-of YYYY-MM-DD
               [--requested-by NAME] [--label "October W4"] [--dry-run] [--json report.json]

Streams handled (pack tab -> workbook tab):
  1_CashIn_OCP    -> 'SRC - Cash In (AX Jul-Aug)'          dedupe: invoice ref + payment ref + amount
  2_CashOut_SAP   -> '<Month> Actual Payments'              dedupe: document no + clearing doc + amount; closes matched ongoing lines
  3_Approved      -> OctOngoing table, AA = Yes             dedupe: normalised ref + supplier (promotes pending -> approved)
  4_Pending       -> OctOngoing table, AA = No              dedupe: normalised ref + supplier
  5a_LC_Paid      -> 'SRC - LC Payments'                    dedupe: date + amount + LC ref; marks LC situation lines paid
  5b_LC_Situation -> new tab 'LC Ongoing (dd-mm)' + repoint  replaces the previous situation
  6_Customs       -> 'DD'                                   dedupe: date + amount + BU code
  7_Standing      -> EXEC E6/E7, PSR reserve, PROC compensable
Every load appends a 'Data Sources' row. The output is saved with fullCalcOnLoad so Excel recalculates on open.
Never saves with data_only; never deletes rows."""
import sys, os, re, json, copy, datetime as dt, argparse
from collections import defaultdict, Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfcommon import (SHEETS, BUS, TIERS, BU_CODES, LC_PROJECT, EXCLUDED_TIERS, canon_bu, norm_ref, norm_supplier,
                      to_date, month_start, num, fmt)

ap = argparse.ArgumentParser()
ap.add_argument('master'); ap.add_argument('pack'); ap.add_argument('--out'); ap.add_argument('--as-of', required=True)
ap.add_argument('--requested-by', default='user'); ap.add_argument('--label', default=''); ap.add_argument('--dry-run', action='store_true')
ap.add_argument('--json'); ap.add_argument('--pack-name')
A = ap.parse_args()
AS_OF = dt.datetime.strptime(A.as_of, '%Y-%m-%d').date()
PACK = A.pack_name or os.path.basename(A.pack)
RUN_AT = dt.datetime.now().strftime('%Y-%m-%d %H:%M')
if not A.dry_run and not A.out: sys.exit('--out is required unless --dry-run')

print(f"Loading master {A.master} …")
wb = load_workbook(A.master)                      # formulas kept
pk = load_workbook(A.pack, data_only=True)
REPORT = {'streams': {}, 'needs_ruling': [], 'warnings': [], 'as_of': A.as_of, 'pack': PACK}
log = lambda s: print('  ' + s)

# ------------------------------------------------------------------ helpers
def pack_rows(tab):
    if tab not in pk.sheetnames: return [], []
    ws = pk[tab]; hdr = [str(c.value).strip() if c.value is not None else '' for c in ws[3]]
    out = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not any(v not in (None, '') for v in row): continue
        out.append({hdr[i]: row[i] for i in range(min(len(hdr), len(row))) if hdr[i]})
    return hdr, out

def H(d, *starts):
    """get a pack value whose header starts with any of the given prefixes"""
    for k, v in d.items():
        for s in starts:
            if k.lower().startswith(s.lower()): return v
    return None

def last_row(ws, cols=('A', 'B', 'C', 'D', 'E'), start=2):
    mr = ws.max_row
    for r in range(mr, start - 1, -1):
        if any(ws[f'{c}{r}'].value not in (None, '') for c in cols): return r
    return start - 1

def copy_formula(ws, src_row, dst_row, col):
    v = ws[f'{col}{src_row}'].value
    if isinstance(v, str) and v.startswith('='):
        f = re.sub(r'(?<![A-Za-z$])(\$?[A-Z]{1,3})' + str(src_row) + r'(?!\d)', lambda m: m.group(1) + str(dst_row), v)
        ws[f'{col}{dst_row}'] = f; return True
    return False

def is_formula(cell): return isinstance(cell.value, str) and cell.value.startswith('=')

def fill(ws, r, col, value, force=False):
    """write a value unless the row above carries a formula in that column (then copy the formula), except when forced"""
    if isinstance(col, int): col = L(col)
    if not force and r > 2 and is_formula(ws[f'{col}{r-1}']):
        copy_formula(ws, r - 1, r, col); return
    ws[f'{col}{r}'] = value

def fill_rest(ws, r, maxcol, written):
    for c in range(1, maxcol + 1):
        if c not in written and r > 2: copy_formula(ws, r - 1, r, L(c))

def style_like(ws, src_row, dst_row, maxcol):
    for c in range(1, maxcol + 1):
        s, d = ws.cell(src_row, c), ws.cell(dst_row, c)
        if s.has_style: d._style = copy.copy(s._style)

OH_NAMES = set()
for row in wb[SHEETS['oh']].iter_rows(min_row=2, values_only=True):
    for v in (row[3] if len(row) > 3 else None, row[4] if len(row) > 4 else None):
        if v: OH_NAMES.add(norm_supplier(v))
TIER_MASTER = {}
for row in wb[SHEETS['tiering']].iter_rows(min_row=5, values_only=True):
    if row[0]: TIER_MASTER[norm_supplier(row[0])] = (str(row[1] or ''), str(row[2] or ''))

def lookup_tier(name, given=None):
    if given not in (None, ''):
        g = str(given).strip(); return g if g in TIERS or g in ('Unclassified', 'Not on OCP Tier') else g
    k = norm_supplier(name)
    if k in TIER_MASTER:
        nat, t = TIER_MASTER[k]; t = t.strip()
        if t in TIERS: return t
        if t in ('1', '2', '3') and nat.lower().startswith('inter'): return f'Intl {t}'
        if t: return t
    return 'Unclassified'

def is_oh(name): return 'OH' if norm_supplier(name) in OH_NAMES else ''

def add_source_row(written_to, source, obj, request, params, rows, changes, notes):
    ws = wb[SHEETS['sources']]; r = last_row(ws, ('A', 'B')) + 1
    vals = [written_to, source, obj, 'Loaded by load_pack.py (Claude, cashflow-bu-update skill)', request, params, RUN_AT,
            A.requested_by, rows, changes, notes]
    for i, v in enumerate(vals, 1): ws.cell(r, i, v)
    style_like(ws, r - 1, r, 11)

# ------------------------------------------------------------------ ongoing table index
og = wb[SHEETS['ongoing']]
tbl = og.tables['OctOngoing']; og_last = int(re.search(r'(\d+)$', tbl.ref).group(1))
og_hdr = {str(og.cell(1, c).value).strip(): c for c in range(1, og.max_column + 1) if og.cell(1, c).value}
def ogc(prefix):
    for k, c in og_hdr.items():
        if k.lower().startswith(prefix.lower()): return c
    raise KeyError(prefix)
C_BU, C_SUP, C_PO, C_REF, C_AMT, C_DUE, C_TIER, C_PRIO, C_COM, C_SRC = (ogc(x) for x in ['BU', 'Supplier', 'PO', 'Reference', 'Amount', 'Due date', 'Tier', 'Priority', 'Comment', 'Source'])
C_STAT, C_KEY, C_OH, C_OVR, C_INLOG, C_ADDED, C_BASIS, C_APPR, C_POS = (ogc(x) for x in ['Sept payment status', 'Ref key', 'OH?', 'Manual status override', 'In payment tracking log', 'Date added', 'Date added —', 'Approved payments (Oct)?', 'Oct position'])
OG_INDEX = defaultdict(list)       # (refkey, supplier) -> rows ; refkey -> rows
for r in range(2, og_last + 1):
    ref = og.cell(r, C_REF).value
    if ref is None and og.cell(r, C_BU).value is None: continue
    kv = og.cell(r, C_KEY).value
    key = norm_ref(ref) if (kv is None or (isinstance(kv, str) and kv.startswith('='))) else norm_ref(kv)
    OG_INDEX[key].append(r)
def og_status(r):
    v = og.cell(r, C_OVR).value
    return str(v if v not in (None, '') else og.cell(r, C_STAT).value or '')

# ================================================================== 1. Cash IN
hdr, rows = pack_rows('1_CashIn_OCP')
if rows:
    ws = wb[SHEETS['cashin']]; lr = last_row(ws, ('I', 'Z', 'AD'))
    seen = set()
    for row in ws.iter_rows(min_row=2, max_row=lr, min_col=1, max_col=31, values_only=True):
        seen.add((norm_ref(row[8]), norm_ref(row[30]), round(num(row[25]), 2)))
    added = skipped = 0; first = lr + 1; tot = 0.0
    for d in rows:
        ref, pay_ref, amt = H(d, 'Réference', 'Reference'), H(d, 'Réf Paiement', 'Ref Paiement'), num(H(d, 'Net à payer'))
        key = (norm_ref(ref), norm_ref(pay_ref), round(amt, 2))
        if key in seen: skipped += 1; continue
        pdate = to_date(H(d, 'Date Paiement')); bu = canon_bu(H(d, 'BU')); flag = H(d, 'Labor')
        if pdate is None or bu is None: REPORT['needs_ruling'].append(f"1_CashIn_OCP: '{ref}' missing payment date or BU — not loaded"); continue
        if flag not in ('Labor', 'Non-Labor'): REPORT['needs_ruling'].append(f"1_CashIn_OCP: '{ref}' Labor/Non-Labor not given — loaded as Unclassified collection"); flag = None
        lr += 1; r = lr
        if not A.dry_run:
            vals = {'C': H(d, 'ID DOC'), 'G': H(d, 'Fournisseur'), 'I': ref, 'R': H(d, 'Commande'), 'U': H(d, 'Division'), 'Z': amt,
                    'AC': H(d, 'Pièce comptable'), 'AD': dt.datetime.combine(pdate, dt.time()), 'AE': pay_ref, 'AI': H(d, 'Filiale'),
                    'AO': bu if bu in BUS else 'Unmapped', 'AP': flag, 'AT': f"Intake pack {PACK} ({AS_OF}) — {H(d, 'Source') or ''}"}
            style_like(ws, r - 1, r, 46)
            for c, v in vals.items(): fill(ws, r, c, v, force=True)
            fill_rest(ws, r, 46, {CI(c) for c in vals})
        seen.add(key); added += 1; tot += amt
    if added and lr > 30086: REPORT['warnings'].append('SRC - Cash In exceeds row 30086 — extend the SUMIFS ranges on the report')
    REPORT['streams']['1_CashIn_OCP'] = {'added': added, 'skipped_duplicates': skipped, 'amount_added': round(tot, 2), 'rows': f'{first}:{lr}' if added else ''}
    log(f"1_CashIn_OCP: +{added} rows ({fmt(tot)} MAD), {skipped} duplicates")
    if added and not A.dry_run:
        add_source_row(f"'{SHEETS['cashin']}'!A{first}:AT{lr} ({added} rows appended)", f"Intake pack {PACK}", '1_CashIn_OCP',
                       'OCP collections transcribed from pack (lines without native OCP file)', 'dedupe invoice ref + payment ref + amount',
                       added, 'BU / Labor flag as given in pack', f"Cash in added {fmt(tot)} MAD; as-of {AS_OF}")

# ================================================================== 2. SAP payments
hdr, rows = pack_rows('2_CashOut_SAP')
if rows:
    bym = defaultdict(list)
    for d in rows:
        cd = to_date(H(d, 'Clearing Date'))
        if cd is None: REPORT['needs_ruling'].append(f"2_CashOut_SAP: doc {H(d,'Document Number')} has no clearing date — not loaded"); continue
        bym[month_start(cd)].append(d)
    for m, lst in bym.items():
        tab = f"{m.strftime('%B')} Actual Payments"
        if tab not in wb.sheetnames:
            REPORT['needs_ruling'].append(f"2_CashOut_SAP: {len(lst)} payments cleared in {m:%B %Y} but tab '{tab}' does not exist — month roll-forward needed first (HOW TO UPDATE §D)"); continue
        ws = wb[tab]; lr = last_row(ws, ('H', 'C', 'G'))
        hdr_ws = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
        def wc(prefix):
            for k, c in hdr_ws.items():
                if k.lower().startswith(prefix.lower()): return c
            return None
        c_amt, c_pos, c_sup, c_name, c_doc, c_dtype, c_ref, c_clr, c_clrd, c_post, c_po, c_type = (wc(x) for x in
            ['Montant', 'Paiement positive', 'Supplier', 'Name', 'Document Number', 'Document Type', 'Reference', 'Clearing Document', 'Clearing Date', 'Posting Date', 'Purchasing Document', 'Type'])
        c_inc, c_tier, c_bu, c_src, c_oh, c_match = (wc(x) for x in ['Include in', 'Tier (report)', 'BU (mapped)', 'BU mapping source', 'OH?', 'Matched ongoing'])
        if None in (c_amt, c_doc, c_clr, c_inc, c_tier, c_bu, c_oh): REPORT['warnings'].append(f"{tab}: helper columns not found — check layout"); continue
        seen = set()
        for row in ws.iter_rows(min_row=2, max_row=lr, values_only=True):
            seen.add((norm_ref(row[c_doc - 1]), norm_ref(row[c_clr - 1]), round(num(row[c_amt - 1]), 2)))
        added = skipped = 0; first = lr + 1; tot = 0.0; closed = []
        for d in lst:
            doc, clr, amt = H(d, 'Document Number'), H(d, 'Clearing Document'), num(H(d, 'Montant'))
            key = (norm_ref(doc), norm_ref(clr), round(amt, 2))
            if key in seen: skipped += 1; continue
            name = H(d, 'Name') or ''; ref = H(d, 'Reference'); typ = str(H(d, 'Type') or ''); pos = str(H(d, 'Paiement positive') or '')
            given = H(d, 'BU'); bu = 'Overhead (corporate)' if str(given or '').lower().startswith('overhead') else canon_bu(given)
            tier = lookup_tier(name, H(d, 'Tier')); oh = is_oh(name); match = None; src = 'invoice lines / pack BU' if bu else ''
            # references to test: the payment's own reference plus the invoice lines of its clearing document
            inv_refs = [x.strip() for x in str(H(d, 'Invoice refs') or '').split(';') if x.strip()]
            inv_amts = [num(x) for x in str(H(d, 'Invoice amounts') or '').split(';') if x.strip()]
            pairs = list(zip(inv_refs, inv_amts)) if len(inv_refs) == len(inv_amts) else [(x, None) for x in inv_refs]
            if ref not in (None, '') and str(ref) not in inv_refs: pairs.append((str(ref), abs(amt) if not pairs else None))
            matches = []
            for rf, ia in pairs:
                for r0 in OG_INDEX.get(norm_ref(rf), []):
                    if not og_status(r0).startswith('Open'): continue
                    oamt = abs(num(og.cell(r0, C_AMT).value)); target = ia if ia is not None else abs(amt)
                    ok = abs(oamt - target) <= max(1.0, 0.01 * oamt)
                    sup_ok = norm_supplier(og.cell(r0, C_SUP).value)[:6] == norm_supplier(name)[:6]
                    if not ok and not sup_ok: continue      # same reference on a different supplier with a different amount: not a match
                    matches.append((r0, rf, oamt, target, ok))
            if matches:
                match = ','.join(str(x[0]) for x in matches)
                if not bu: bu = og.cell(matches[0][0], C_BU).value; src = f'Ongoing row(s) {match} (reference match)'
                for r0, rf, oamt, target, ok in matches:
                    if ok:
                        closed.append((r0, rf, target))
                        if not A.dry_run:
                            og.cell(r0, C_OVR, 'Paid in Oct' if m.month == 10 else f'Paid in {m:%b}')
                            og.cell(r0, C_COM, f"{og.cell(r0, C_COM).value or ''} | Paid {to_date(H(d,'Clearing Date'))} SAP doc {doc} clearing {clr} (pack {PACK})".strip(' |'))
                    else:
                        REPORT['needs_ruling'].append(f"2_CashOut_SAP: invoice '{rf}' ({name}) paid {fmt(target)} on doc {doc} vs ongoing row {r0} amount {fmt(oamt)} — partial/short payment: confirm closing and balance row")
            if tier in EXCLUDED_TIERS: inc = 'No - excluded tier'
            elif not typ.lower().startswith('paiement'): inc = 'No - invoice line'
            elif pos and not pos.lower().startswith('oui'): inc = 'No - payment reversal line'
            elif bu == 'Overhead (corporate)': inc = 'No - corporate OH'; oh = 'OH'
            else: inc = 'Yes'
            if not bu:
                bu = '(unmapped)'; src = 'No match in ongoing position or payment history'
                if inc == 'Yes': REPORT['needs_ruling'].append(f"2_CashOut_SAP: payment {fmt(amt)} {name} doc {doc} (invoices {', '.join(inv_refs[:3]) or 'n/a'}) — no BU found: give a BU (counted in consolidated Unclassified until then)")
            note = str(H(d, 'Source') or '')
            if 'mixed BUs' in note and inc == 'Yes': REPORT['needs_ruling'].append(f"2_CashOut_SAP: payment {fmt(amt)} {name} doc {doc} — {note.split(';')[-2].strip() if ';' in note else note}")
            if tier == 'Unclassified' and inc == 'Yes': REPORT['needs_ruling'].append(f"2_CashOut_SAP: supplier '{name}' has no tier on Tiering Master — tier needed")
            lr += 1; r = lr
            if not A.dry_run:
                style_like(ws, r - 1, r, ws.max_column)
                for c, v in [(c_type, typ), (c_amt, amt), (c_pos, pos or None), (c_sup, H(d, 'Supplier')), (c_name, name), (c_doc, doc), (c_dtype, H(d, 'Document Type')),
                             (c_ref, ref), (c_clr, clr), (c_clrd, dt.datetime.combine(to_date(H(d, 'Clearing Date')), dt.time())), (c_post, to_date(H(d, 'Posting Date')) and dt.datetime.combine(to_date(H(d, 'Posting Date')), dt.time())),
                             (c_po, H(d, 'Purchasing Document')), (c_inc, inc), (c_tier, tier), (c_bu, bu), (c_src, f"{src} — {str(H(d, 'Source') or '')[:150]}"), (c_oh, oh), (c_match, match)]:
                    if c: ws.cell(r, c, v)
            seen.add(key); added += 1
            if inc == 'Yes': tot += amt
        if added and lr > 2000: REPORT['warnings'].append(f"{tab} exceeds row 2000 — extend the SUMIFS ranges on the report (column E)")
        REPORT['streams'][f'2_CashOut_SAP ({tab})'] = {'added': added, 'skipped_duplicates': skipped, 'included_amount_added': round(tot, 2), 'ongoing_lines_closed': len(closed), 'rows': f'{first}:{lr}' if added else ''}
        log(f"2_CashOut_SAP -> {tab}: +{added} rows (included {fmt(tot)} MAD), {skipped} duplicates, {len(closed)} ongoing lines closed")
        if added and not A.dry_run:
            add_source_row(f"'{tab}'!A{first}:{L(ws.max_column)}{lr} ({added} rows appended); OctOngoing override set on {len(closed)} rows", f"Intake pack {PACK}", '2_CashOut_SAP',
                           'SAP payment lines cleared, transcribed from pack', 'dedupe document no + clearing doc + amount; ongoing match on normalised reference, amount within 0.5%',
                           added, f"Closed ongoing rows: {', '.join(str(c[0]) for c in closed) or 'none'}", f"Included cash out added {fmt(-tot)} MAD; as-of {AS_OF}")

# ================================================================== 3/4. Ongoing (approved / pending)
def load_ongoing(tab, approved):
    global og_last
    hdr, rows = pack_rows(tab)
    if not rows: return
    added = skipped = promoted = 0; first = og_last + 1; tot = 0.0; label = A.label or f'pack {AS_OF:%d/%m}'
    for d in rows:
        bu = canon_bu(H(d, 'BU')); sup = H(d, 'Supplier'); ref = H(d, 'Reference'); amt = num(H(d, 'Amount'))
        if not bu or bu not in BUS or not sup or ref in (None, ''):
            REPORT['needs_ruling'].append(f"{tab}: line '{sup} / {ref}' missing BU/supplier/reference — not loaded"); continue
        if amt > 0: REPORT['warnings'].append(f"{tab}: '{ref}' ({sup}) positive amount kept as a credit/reversal line (+{amt:,.2f})")
        key = norm_ref(ref); existing = OG_INDEX.get(key, [])
        same = [r for r in existing if norm_supplier(og.cell(r, C_SUP).value) == norm_supplier(sup) or abs(num(og.cell(r, C_AMT).value) - amt) <= 1]
        if same:
            r0 = same[0]; st = og_status(r0)
            if st.startswith('Open'):
                if approved and og.cell(r0, C_APPR).value != 'Yes':
                    if not A.dry_run:
                        og.cell(r0, C_INLOG, 'Yes'); og.cell(r0, ogc('Check vs tracking log'), f"Approved — {H(d, 'Approved via') or 'user list'} {AS_OF:%d/%m/%Y}; treated as in tracking log per user")
                        og.cell(r0, C_COM, f"{og.cell(r0, C_COM).value or ''} | Approved {AS_OF} via {H(d, 'Approved via') or 'pack'} ({H(d, 'Source') or PACK})".strip(' |'))
                    promoted += 1
                else: skipped += 1
            else:
                skipped += 1; REPORT['warnings'].append(f"{tab}: '{ref}' ({sup}) already on the ongoing position with status '{st}' — not re-added")
            continue
        og_last += 1; r = og_last; tier = lookup_tier(sup, H(d, 'Tier'))
        if tier == 'Unclassified': REPORT['needs_ruling'].append(f"{tab}: supplier '{sup}' not on Tiering Master — tier needed (loaded as Unclassified)")
        if not A.dry_run:
            style_like(og, r - 1, r, og.max_column)
            for c in range(1, og.max_column + 1): og.cell(r, c).value = None
            due = to_date(H(d, 'Due date'))
            via = H(d, 'Approved via') or H(d, 'Requested by') or 'pack'
            label = str(H(d, 'Status label') or '').strip() or (A.label or f'pack {AS_OF:%d/%m}')
            vals = {C_BU: bu, C_SUP: sup, C_PO: H(d, 'PO'), C_REF: ref, C_AMT: amt, C_DUE: dt.datetime.combine(due, dt.time()) if due else None, C_TIER: tier,
                    C_PRIO: H(d, 'Priority'), C_COM: H(d, 'Comment'), C_SRC: f"{H(d, 'Source') or ''} (pack {PACK})".strip(),
                    C_OVR: f'Open - {label}',
                    # approved = counted as "in tracking log" (user-approved lists are treated as in the log, standing rule)
                    C_INLOG: 'Yes' if approved else ('Yes' if str(H(d, 'In tracking log') or '').lower().startswith('y') else 'No'),
                    ogc('Tracking log — detail'): f"{'In tracking log / user-approved' if approved else 'Not in tracking log'} (pack {PACK})",
                    ogc('Check vs tracking log'): (f"Approved — {via} {AS_OF:%d/%m/%Y}; treated as in tracking log per user" if approved else f"Pending payment approval — {via} {AS_OF:%d/%m/%Y}; not in tracking log"),
                    C_ADDED: dt.datetime.combine(AS_OF, dt.time()), C_BASIS: f'Intake pack {PACK}'}
            for c, v in vals.items(): og.cell(r, c, v)
            # status (N), ref key (O), OH? (P), approved (AA) and position (AB) are table formulas: copy them down
            fill_rest(og, r, og.max_column, set(vals))
        OG_INDEX[key].append(r); added += 1; tot += amt
    if not A.dry_run and added: tbl.ref = f"A1:{L(og.max_column)}{og_last}"
    REPORT['streams'][tab] = {'added': added, 'promoted_pending_to_approved': promoted, 'skipped_duplicates': skipped, 'amount_added': round(tot, 2), 'rows': f'{first}:{og_last}' if added else ''}
    log(f"{tab}: +{added} lines ({fmt(tot)} MAD), {promoted} promoted to approved, {skipped} duplicates")
    if (added or promoted) and not A.dry_run:
        add_source_row(f"'{SHEETS['ongoing']}'!A{first}:AH{og_last} ({added} rows in OctOngoing, table resized); {promoted} rows AA set to Yes", f"Intake pack {PACK}", tab,
                       ('Approved' if approved else 'Pending-approval') + ' commitments transcribed from pack (emails / lists / PDFs)',
                       'dedupe normalised reference + supplier; tier from Tiering Master; OH from OH Suppliers', added,
                       f"{promoted} existing pending lines promoted to approved", f"{'Approved' if approved else 'Pending'} ongoing added {fmt(tot)} MAD; as-of {AS_OF}")
load_ongoing('3_Approved', True)
load_ongoing('4_Pending', False)

# ================================================================== 3b. Tracking-log check on open ongoing lines
hdr, rows = pack_rows('3b_LogCheck')
if rows:
    stamp = f"{AS_OF:%d/%m}"; c_st = og.max_column + 1; c_dt = og.max_column + 2
    if not A.dry_run:
        og.cell(1, c_st, f"Tracking log check {stamp} — status"); og.cell(1, c_dt, f"Tracking log check {stamp} — detail")
        for c in (c_st, c_dt): og.cell(1, c)._style = copy.copy(og.cell(1, og.max_column - 2)._style)
        from openpyxl.worksheet.table import TableColumn
        ids = [tc.id for tc in tbl.tableColumns]
        tbl.tableColumns.append(TableColumn(id=max(ids) + 1, name=f"Tracking log check {stamp} — status"))
        tbl.tableColumns.append(TableColumn(id=max(ids) + 2, name=f"Tracking log check {stamp} — detail"))
        tbl.ref = f"A1:{L(c_dt)}{og_last}"
    flagged = promoted = 0; paid_in_log = 0; st_count = Counter()
    for d in rows:
        r = int(num(H(d, 'Ongoing row'))); inlog = str(H(d, 'In log') or '').lower().startswith('y'); status = H(d, 'Overall status'); st_count[str(status)] += 1 if inlog else 0
        detail = ' | '.join(str(x) for x in [H(d, 'AP blocage'), H(d, 'Treasury comment'), H(d, 'Treasury blocage'), f"paid {to_date(H(d, 'Payment date'))}" if H(d, 'Payment date') else None, f"log amt {H(d, 'Log amount')}" if H(d, 'Log amount') else None, H(d, 'Match basis')] if x not in (None, '', 'None'))
        if not A.dry_run:
            og.cell(r, c_st, (f"In log — {status}" if inlog else "Not in log") + (" — kept Open (await SAP)" if inlog and str(status).upper() == 'PAID' else ''))
            og.cell(r, c_dt, detail or None)
            if inlog:
                if og.cell(r, C_INLOG).value != 'Yes': promoted += 1
                og.cell(r, C_INLOG, 'Yes'); og.cell(r, ogc('Tracking log — Overall'), status); og.cell(r, ogc('Tracking log — detail'), detail or None)
        if inlog: flagged += 1
        if inlog and str(status).upper() == 'PAID': paid_in_log += 1
    REPORT['streams']['3b_LogCheck'] = {'open_lines_checked': len(rows), 'in_log': flagged, 'newly_approved_via_log': promoted, 'paid_in_log_kept_open': paid_in_log, 'log_status_counts': dict(st_count)}
    log(f"3b_LogCheck: {len(rows)} open lines checked, {flagged} in log ({promoted} newly approved), {paid_in_log} PAID in log kept Open")
    if not A.dry_run:
        add_source_row(f"'{SHEETS['ongoing']}'!{L(c_st)}1:{L(c_dt)}{og_last} (new check columns) + U/V/W on {flagged} rows", f"Intake pack {PACK} (3b_LogCheck from Payment tracking log)", '3b_LogCheck',
                       'Tracking-log status refreshed for every open ongoing line', 'match on normalised reference, then supplier/amount; PAID lines kept Open (standing rule)', len(rows),
                       f"{promoted} lines moved pending -> approved (now in log)", f"as-of {AS_OF}")

# ================================================================== 8. Rulings (user decisions applied to existing lines)
# tab 8_Rulings columns: Action | Target | Value | Reason | Source
#   close   | ongoing row      | (paid amount)            -> Manual status override = Paid in Oct
#   balance | ongoing row      | balance amount (negative)-> new OctOngoing row '<ref> - balance', same BU/supplier/tier, Open, same approval
#   include | SAP document no. | Yes / No - corporate OH / No - excluded tier -> October tab AF (+ AH 'Overhead (corporate)', AJ 'OH' for corporate)
#   bu      | SAP document no. | canonical BU             -> October tab AH
hdr, rows = pack_rows('8_Rulings')
if rows:
    oct_ws = wb[SHEETS['pay_oct']]; oh_ws = {str(oct_ws.cell(1, c).value).strip(): c for c in range(1, oct_ws.max_column + 1) if oct_ws.cell(1, c).value}
    def occ(prefix):
        for k, c in oh_ws.items():
            if k.lower().startswith(prefix.lower()): return c
    c_doc, c_inc, c_bu, c_oh, c_src = occ('Document Number'), occ('Include in'), occ('BU (mapped)'), occ('OH?'), occ('BU mapping source')
    if 'C_TIER' not in globals(): C_TIER = ogc('Tier')
    done = Counter(); bal_rows = []
    for d in rows:
        act = str(H(d, 'Action') or '').strip().lower(); tgt = H(d, 'Target'); val = H(d, 'Value'); why = H(d, 'Reason') or ''
        if act == 'close':
            r = int(num(tgt))
            if not A.dry_run:
                og.cell(r, C_OVR, 'Paid in Oct'); og.cell(r, C_COM, f"{og.cell(r, C_COM).value or ''} | Closed per user ruling {AS_OF}: {why}".strip(' |'))
            done['close'] += 1
        elif act == 'balance':
            r = int(num(tgt)); amt = -abs(num(val)); og_last += 1; nr = og_last
            if not A.dry_run:
                style_like(og, nr - 1, nr, og.max_column)
                for c in range(1, og.max_column + 1): og.cell(nr, c).value = None
                vals = {C_BU: og.cell(r, C_BU).value, C_SUP: og.cell(r, C_SUP).value, C_PO: og.cell(r, C_PO).value, C_REF: f"{og.cell(r, C_REF).value} - balance", C_AMT: amt,
                        C_DUE: og.cell(r, C_DUE).value, C_TIER: og.cell(r, C_TIER).value, C_PRIO: og.cell(r, C_PRIO).value, C_COM: f"Balance after short payment of row {r} ({why}); user ruling {AS_OF}",
                        C_SRC: og.cell(r, C_SRC).value, C_OVR: f"Open - balance {AS_OF:%d/%m}", C_INLOG: ('Yes' if str(H(d, 'Approval') or '').lower().startswith('y') else 'No') if H(d, 'Approval') not in (None, '') else og.cell(r, C_INLOG).value,
                        ogc('Check vs tracking log'): og.cell(r, ogc('Check vs tracking log')).value, C_ADDED: dt.datetime.combine(AS_OF, dt.time()), C_BASIS: f'Balance row (ruling {AS_OF})'}
                for c, v in vals.items(): og.cell(nr, c, v)
                fill_rest(og, nr, og.max_column, set(vals))
            bal_rows.append(nr); done['balance'] += 1
        elif act == 'pending':
            r = int(num(tgt))
            if not A.dry_run:
                og.cell(r, C_INLOG, 'No'); og.cell(r, ogc('Check vs tracking log'), f"Pending payment approval — user ruling {AS_OF}: {why}")
                og.cell(r, C_COM, f"{og.cell(r, C_COM).value or ''} | Set to pending per user ruling {AS_OF}".strip(' |'))
            done['pending'] += 1
        elif act == 'tier':          # Target = ongoing row (number) or SAP document number (string starting with 'doc:')
            if str(tgt).startswith('doc:'):
                docno = str(tgt)[4:].strip(); c_tier = occ('Tier (report)')
                hit = [rr for rr in range(2, oct_ws.max_row + 1) if str(oct_ws.cell(rr, c_doc).value) == docno]
                if not A.dry_run:
                    for rr in hit: oct_ws.cell(rr, c_tier, val); oct_ws.cell(rr, c_src, f"{oct_ws.cell(rr, c_src).value or ''} | tier {val} per user ruling {AS_OF}: {why}".strip(' |'))
                done['tier'] += len(hit)
            else:
                r = int(num(tgt))
                if not A.dry_run:
                    og.cell(r, C_TIER, val); og.cell(r, C_COM, f"{og.cell(r, C_COM).value or ''} | tier {val} per user ruling {AS_OF}: {why}".strip(' |'))
                done['tier'] += 1
        elif act == 'master':        # Target = supplier name as written on the lines; Value = 'National|International|<tier>'
            tmw = wb[SHEETS['tiering']]; nat, tier = [x.strip() for x in str(val).split('|')][:2]
            rr = tmw.max_row + 1
            while tmw.cell(rr - 1, 1).value in (None, '') and rr > 6: rr -= 1
            if not A.dry_run:
                tmw.cell(rr, 1, str(tgt)); tmw.cell(rr, 2, nat); tmw.cell(rr, 3, int(tier) if tier.isdigit() else tier); tmw.cell(rr, 4, 'Research'); tmw.cell(rr, 5, f"{why} (user ruling {AS_OF})")
            done['master'] += 1
        elif act == 'remove':        # Target = ongoing row; Value = status text (e.g. 'Removed - per user dd/mm (...)')
            r = int(num(tgt))
            if not A.dry_run:
                og.cell(r, C_OVR, str(val) or f'Removed - per user {AS_OF:%d/%m}'); og.cell(r, C_COM, f"{og.cell(r, C_COM).value or ''} | {why}".strip(' |'))
            done['remove'] += 1
        elif act == 'master_remove':  # Target = supplier name exactly as on the Tiering Master
            tmw = wb[SHEETS['tiering']]; hit = [rr for rr in range(5, tmw.max_row + 1) if str(tmw.cell(rr, 1).value).strip().upper() == str(tgt).strip().upper()]
            if not A.dry_run:
                for rr in hit:
                    for c in range(1, 6): tmw.cell(rr, c).value = None
            done['master_remove'] += len(hit)
        elif act in ('include', 'bu'):
            hit = [rr for rr in range(2, oct_ws.max_row + 1) if str(oct_ws.cell(rr, c_doc).value) == str(tgt).strip() and str(oct_ws.cell(rr, c_inc).value) != 'No - invoice line']
            for rr in hit:
                if not A.dry_run:
                    if act == 'include':
                        oct_ws.cell(rr, c_inc, val)
                        if 'corporate' in str(val).lower(): oct_ws.cell(rr, c_bu, 'Overhead (corporate)'); oct_ws.cell(rr, c_oh, 'OH')
                    else: oct_ws.cell(rr, c_bu, canon_bu(val))
                    oct_ws.cell(rr, c_src, f"{oct_ws.cell(rr, c_src).value or ''} | user ruling {AS_OF}: {why}".strip(' |'))
            done[act] += len(hit)
            if not hit: REPORT['warnings'].append(f"8_Rulings: SAP document {tgt} not found on '{SHEETS['pay_oct']}'")
        else: REPORT['warnings'].append(f"8_Rulings: unknown action '{act}'")
    if not A.dry_run and bal_rows: tbl.ref = f"A1:{L(og.max_column)}{og_last}"
    REPORT['streams']['8_Rulings'] = dict(done); log(f"8_Rulings: {dict(done)}")
    if not A.dry_run and done:
        add_source_row(f"OctOngoing rows closed/added ({dict(done)}); balance rows {bal_rows or 'none'}; '{SHEETS['pay_oct']}' AF/AH/AJ on ruled documents", f"User rulings {AS_OF} (pack {PACK}, tab 8_Rulings)", '8_Rulings',
                       'User decisions on short/over-payments, corporate charges and BU mapping', 'see Reason column of the pack', len(rows), 'None', f"as-of {AS_OF}")

# ================================================================== 5a. LC paid
hdr, rows = pack_rows('5a_LC_Paid')
lc_paid_refs = set()
if rows:
    ws = wb[SHEETS['lc']]; lr = last_row(ws, ('E', 'G'))
    seen = set()
    for row in ws.iter_rows(min_row=2, max_row=lr, max_col=13, values_only=True):
        seen.add((to_date(row[4]), round(num(row[6]), 2), norm_ref(row[12])))
    added = skipped = 0; first = lr + 1; tot = 0.0
    for d in rows:
        date = to_date(H(d, 'Date')); amt = num(H(d, 'Débit', 'Debit')); lcref = H(d, 'Ref LC'); proj = str(H(d, 'Project') or '').strip().upper()
        if date is None or not amt: REPORT['needs_ruling'].append(f"5a_LC_Paid: line '{lcref}' missing date/amount — not loaded"); continue
        key = (date, round(amt, 2), norm_ref(lcref))
        if key in seen: skipped += 1; continue
        if proj not in BU_CODES: REPORT['needs_ruling'].append(f"5a_LC_Paid: '{lcref}' project code '{proj}' unknown — not attributed to a BU (loaded, but outside BU blocks)")
        lr += 1; r = lr
        if not A.dry_run:
            style_like(ws, r - 1, r, 37)
            vals = {'B': 'JESA', 'E': dt.datetime.combine(date, dt.time()), 'F': H(d, 'Libellé'), 'G': amt, 'I': H(d, 'MNT DEVISE'), 'J': H(d, 'D.Curr'), 'K': H(d, 'PO'),
                    'M': lcref, 'N': H(d, 'Nom fournisseur'), 'S': proj, 'X': H(d, 'Nom fournisseur'), 'Y': BU_CODES.get(proj), 'AE': date.strftime('%B'),
                    'AF': lookup_tier(H(d, 'Nom fournisseur'), H(d, 'Tier')), 'AG': f"Intake pack {PACK} — {H(d, 'Source') or ''}"}
            for c, v in vals.items(): fill(ws, r, c, v, force=c not in ('AF', 'AG'))
            fill_rest(ws, r, 37, {CI(c) for c in vals})
        seen.add(key); added += 1; tot += amt; lc_paid_refs.add(norm_ref(lcref))
    REPORT['streams']['5a_LC_Paid'] = {'added': added, 'skipped_duplicates': skipped, 'amount_added': round(tot, 2), 'rows': f'{first}:{lr}' if added else ''}
    log(f"5a_LC_Paid: +{added} lines ({fmt(tot)} MAD), {skipped} duplicates")
    # mark LC situation lines paid (ref match)
    lo = wb[SHEETS['lc_ongoing']]; marked = []
    if lc_paid_refs and not A.dry_run:
        for rr in range(5, lo.max_row + 1):
            if norm_ref(lo.cell(rr, 5).value) in lc_paid_refs and str(lo.cell(rr, 18).value) != 'Yes':
                lo.cell(rr, 18, 'Yes')
                if not (isinstance(lo.cell(rr, 17).value, str) and str(lo.cell(rr, 17).value).startswith('=')): lo.cell(rr, 17, 0)
                marked.append(rr)
    if added and not A.dry_run:
        add_source_row(f"'{SHEETS['lc']}'!A{first}:AK{lr} ({added} rows appended); '{SHEETS['lc_ongoing']}' R set Yes on rows {marked or 'none'}", f"Intake pack {PACK}", '5a_LC_Paid',
                       'LC bank-statement payments transcribed from pack', 'dedupe date + amount + LC ref; LC situation matched on LC/CREDOC ref', added,
                       f"{len(marked)} LC situation lines set to paid", f"LC paid added {fmt(tot)} MAD; as-of {AS_OF}")

# ================================================================== 5b. LC situation (new dated tab + repoint)
hdr, rows = pack_rows('5b_LC_Situation')
if rows:
    sit = to_date(H(rows[0], 'Situation date')) or AS_OF
    new_tab = f"LC Ongoing ({sit:%d-%m})"
    old_tab = SHEETS['lc_ongoing']
    if new_tab in wb.sheetnames: REPORT['warnings'].append(f"5b_LC_Situation: tab '{new_tab}' already exists — not rebuilt"); rows = []
    if rows:
        if not A.dry_run:
            old = wb[old_tab]; ws = wb.copy_worksheet(old); ws.title = new_tab
            for r in range(5, ws.max_row + 1):
                for c in range(1, 19): ws.cell(r, c).value = None
            ws['A1'] = f"LC MATURITIES — {sit:%B %Y} ongoing (approved) — bank LC situation {sit:%d/%m/%Y}"
            ws['A2'] = f"Source: intake pack {PACK} (sheet 5b_LC_Situation) — maturities of the current month per the bank 'LC avec echéance' situation"
            ws['A3'] = f"Loaded {RUN_AT}. Column R = already paid this month (LC ref found in 'SRC - LC Payments' for the month or given in the pack); Q = 0 when R = Yes, else −L."
        r = 4; tot = 0.0; by_bu = defaultdict(float)
        for d in rows:
            r += 1; proj = str(H(d, 'Project code') or '').strip().upper(); bu = canon_bu(H(d, 'Business Unit')) or BU_CODES.get(proj)
            mat = num(H(d, 'Current-month maturity', 'Oct-26 LC maturity', 'maturity'))
            paid = str(H(d, 'Already paid') or '').lower().startswith('y') or norm_ref(H(d, 'LC / CREDOC')) in lc_paid_refs
            if not bu: REPORT['needs_ruling'].append(f"5b_LC_Situation: LC '{H(d, 'LC / CREDOC')}' project '{proj}' — BU unknown, excluded from ongoing")
            q = 0 if paid or not bu else -mat; tot += q
            if bu: by_bu[bu] += q
            if not A.dry_run:
                vals = {1: H(d, 'Source row'), 2: H(d, 'L.E.'), 3: proj, 4: H(d, 'PO'), 5: H(d, 'LC / CREDOC'), 6: H(d, 'Supplier'), 7: to_date(H(d, 'LC expiry')) and dt.datetime.combine(to_date(H(d, 'LC expiry')), dt.time()),
                        8: H(d, 'Ccy'), 9: num(H(d, 'LC face')), 10: num(H(d, 'Encours LC (FX)')), 11: num(H(d, 'Encours LC (MAD)')), 12: mat, 13: H(d, 'Action'), 14: H(d, 'Comments'),
                        15: 'pack', 16: bu, 17: f'=IF(R{r}="Yes",0,-N(L{r}))', 18: 'Yes' if paid else 'No'}
                for c, v in vals.items(): ws.cell(r, c, v)
        last = r
        if not A.dry_run:
            ws.cell(last + 2, 1, 'TOTAL'); ws.cell(last + 2, 12, f'=SUM(L5:L{last})'); ws.cell(last + 2, 17, f'=SUM(Q5:Q{last})')
            ws.cell(last + 4, 16, 'Business Unit'); ws.cell(last + 4, 17, f'{sit:%b-%y} LC maturity (MAD)'); ws.cell(last + 4, 18, 'Ongoing included (MAD)')
            for i, b in enumerate(sorted(by_bu), 1):
                ws.cell(last + 4 + i, 16, b); ws.cell(last + 4 + i, 17, f'=SUMIFS(L5:L{last},P5:P{last},P{last+4+i})'); ws.cell(last + 4 + i, 18, f'=SUMIFS(Q5:Q{last},P5:P{last},P{last+4+i})')
            # repoint every formula that references the old situation tab
            pat = re.compile(r"'" + re.escape(old_tab) + r"'!\$([A-Z]+)\$5:\$([A-Z]+)\$\d+")
            n = 0
            for s in wb.worksheets:
                if s.title in (old_tab, new_tab): continue
                for row in s.iter_rows():
                    for cell in row:
                        v = cell.value
                        if isinstance(v, str) and old_tab in v and v.startswith('='):
                            nv = pat.sub(lambda m: f"'{new_tab}'!${m.group(1)}$5:${m.group(2)}${last}", v).replace(f"'{old_tab}'", f"'{new_tab}'")
                            if nv != v: cell.value = nv; n += 1
            REPORT['streams']['5b_LC_Situation'] = {'new_tab': new_tab, 'lines': last - 4, 'ongoing_total': round(tot, 2), 'by_bu': {k: round(v, 2) for k, v in by_bu.items()}, 'formulas_repointed': n}
            add_source_row(f"New tab '{new_tab}' A4:R{last} ({last-4} LCs) + totals; {n} formulas repointed from '{old_tab}'", f"Intake pack {PACK}", '5b_LC_Situation',
                           'Bank LC situation (LC avec echéance) transcribed from pack', 'ongoing = −maturity unless already paid (ref match)', last - 4,
                           f"Supersedes '{old_tab}' for the LC ongoing rows (F) of the report", f"LC ongoing now {fmt(tot)} MAD; as-of {AS_OF}")
        else:
            REPORT['streams']['5b_LC_Situation'] = {'new_tab': new_tab, 'lines': len(rows), 'ongoing_total': round(tot, 2), 'by_bu': {k: round(v, 2) for k, v in by_bu.items()}}
        log(f"5b_LC_Situation: {len(rows)} LCs -> {new_tab}, ongoing {fmt(tot)} MAD")

# ================================================================== 6. Customs
hdr, rows = pack_rows('6_Customs')
if rows:
    ws = wb[SHEETS['customs']]; lr = last_row(ws, ('B', 'E'))
    seen = set()
    for row in ws.iter_rows(min_row=2, max_row=lr, max_col=8, values_only=True):
        seen.add((to_date(row[1]), round(num(row[4]), 2), str(row[7] or '').strip().upper()))
    added = skipped = 0; first = lr + 1; tot = 0.0
    for d in rows:
        date = to_date(H(d, 'Date')); amt = num(H(d, 'Montant sur relevé')); code = str(H(d, 'BU code') or '').strip().upper()
        if date is None or not amt or not code: REPORT['needs_ruling'].append(f"6_Customs: line '{H(d, 'Libellé')}' missing date/amount/BU code — not loaded"); continue
        key = (date, round(amt, 2), code)
        if key in seen: skipped += 1; continue
        if code not in BU_CODES: REPORT['needs_ruling'].append(f"6_Customs: BU code '{code}' unknown — line loaded but not attributed to a BU")
        lr += 1; r = lr
        if not A.dry_run:
            style_like(ws, r - 1, r, 13)
            vals = {'A': H(d, 'Company code') or 'JESA', 'B': dt.datetime.combine(date, dt.time()), 'C': to_date(H(d, 'Valeur')) and dt.datetime.combine(to_date(H(d, 'Valeur')), dt.time()),
                    'D': H(d, 'Libellé'), 'E': amt, 'F': num(H(d, 'Montant du ticket')) or None, 'H': code}
            for c, v in vals.items(): fill(ws, r, c, v, force=True)
            fill_rest(ws, r, 13, {CI(c) for c in vals})
            if not is_formula(ws[f'I{r}']): ws[f'I{r}'] = BU_CODES.get(code, code)
            if not is_formula(ws[f'J{r}']): ws[f'J{r}'] = dt.datetime(date.year, date.month, 1)
        seen.add(key); added += 1; tot += amt
    if added and lr > 2000: REPORT['warnings'].append('DD exceeds row 2000 — extend the customs SUMIFS ranges on the report')
    REPORT['streams']['6_Customs'] = {'added': added, 'skipped_duplicates': skipped, 'amount_added': round(tot, 2), 'rows': f'{first}:{lr}' if added else ''}
    log(f"6_Customs: +{added} lines ({fmt(tot)} MAD), {skipped} duplicates")
    if added and not A.dry_run:
        add_source_row(f"'DD'!A{first}:J{lr} ({added} rows appended)", f"Intake pack {PACK}", '6_Customs', 'Treasury customs detail by BU and date transcribed from pack',
                       'dedupe date + amount + BU code; BU (report) and Month by formula', added, 'None', f"Customs added {fmt(tot)} MAD; as-of {AS_OF}")

def set_as_of(d):
    ex = wb[SHEETS['exec']]; ex['E7'] = dt.datetime.combine(d, dt.time())
    if not is_formula(ex['F7']): ex['F7'] = dt.datetime.combine(d, dt.time())

# ================================================================== 7. Standing
hdr, rows = pack_rows('7_Standing')
if rows:
    ex = wb[SHEETS['exec']]; pc = wb[SHEETS['proc_comp']]; ps = wb[SHEETS['psr_res']]; applied = []
    for d in rows:
        k = str(H(d, 'Input') or '').lower(); v = H(d, 'Value')
        if v in (None, ''): continue
        if k.startswith('master switch'):
            if not A.dry_run: ex['F6' if ex['E6'].value is None else 'E6'] = bool(str(v).strip().lower() in ('true', '1', 'yes', 'oui'))
            applied.append(f'master switch = {v}')
        elif k.startswith('data as of'):
            if not A.dry_run: set_as_of(to_date(v))
            applied.append(f'data as of = {to_date(v)}')
        elif k.startswith('psr reserve — total') or k.startswith('psr reserve - total'):
            if not A.dry_run: ps['B5'] = num(v)
            applied.append(f'PSR reserve total = {fmt(num(v))}')
        elif k.startswith('psr reserve — number') or k.startswith('psr reserve - number'):
            if not A.dry_run: ps['B6'] = int(num(v))
            applied.append(f'PSR months = {int(num(v))}')
        elif k.startswith('proc compensable'):
            bu = canon_bu(k.split('—')[-1].split('-')[-1].strip()) if ('—' in k or '-' in k) else None
            hit = [r for r in range(5, 15) if canon_bu(pc.cell(r, 1).value) == bu]
            if hit:
                if not A.dry_run and not str(pc.cell(hit[0], 3).value).startswith('='): pc.cell(hit[0], 3, num(v))
                applied.append(f'PROC compensable {bu} = {fmt(num(v))}' + (' (formula cell — edit invoice list instead)' if str(pc.cell(hit[0], 3).value).startswith('=') else ''))
            else: REPORT['needs_ruling'].append(f"7_Standing: '{H(d, 'Input')}' — BU not recognised")
        else:
            REPORT['needs_ruling'].append(f"7_Standing: '{H(d, 'Input')}' = {v} must be applied manually (revenue tab / roll-forward)")
    REPORT['streams']['7_Standing'] = {'applied': applied}
    log(f"7_Standing: {', '.join(applied) or 'nothing applied'}")
    if applied and not A.dry_run:
        add_source_row("'EXEC Summary'!E6:E7 / 'PSR Reserve (provision)'!B5:B6 / 'PROC Compensable (standing)'!C5:C14", f"Intake pack {PACK}", '7_Standing',
                       'Standing inputs changed on user instruction', '; '.join(applied), len(applied), 'None', f"as-of {AS_OF}")

# ================================================================== cover + save
if not A.dry_run:
    ex = wb[SHEETS['exec']]
    if not any(str(H(d, 'Input') or '').lower().startswith('data as of') for d in pack_rows('7_Standing')[1]):
        set_as_of(AS_OF)
    wb.calculation.fullCalcOnLoad = True
    print(f"Saving {A.out} …"); wb.save(A.out)
REPORT['dry_run'] = A.dry_run
if A.json: json.dump(REPORT, open(A.json, 'w'), indent=1, default=str, ensure_ascii=False)
print("\nSUMMARY"); [print(f"  {k}: {v}") for k, v in REPORT['streams'].items()]
if REPORT['warnings']: print("\nWARNINGS"); [print('  - ' + w) for w in REPORT['warnings']]
if REPORT['needs_ruling']: print("\nNEEDS A RULING FROM THE USER"); [print('  - ' + w) for w in REPORT['needs_ruling']]
print("\nNext: run shadow_calc.py on the output, explain every headline move by stream, update state/refresh-state.json and state/refresh-log.md.")
