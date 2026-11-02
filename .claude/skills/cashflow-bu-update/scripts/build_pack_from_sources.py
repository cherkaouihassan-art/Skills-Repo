#!/usr/bin/env python3
"""Build an intake pack from the native Drive files (no manual transcription):
  --ocp  <OCP Payments Weekly.xlsx>      sheet 'Payments Key'  -> 1_CashIn_OCP   (invoices paid >= --from, not yet in master)
  --sap  <PAYMENTS ... .xlsx>            sheet 'PAYMENTS OCT'  -> 2_CashOut_SAP  (paiement lines only; loader dedupes)
  --log  <Payment tracking log.xlsx>     sheet 'Invoice Tracker' -> 3b_LogCheck  (status per open ongoing line) + report of log-open lines not on the position
Usage: build_pack_from_sources.py <master.xlsx> --template <CF_INPUT_PACK.xlsx> --out <pack.xlsx> [--ocp f] [--sap f] [--log f] [--from 2026-10-01] [--json report.json]"""
import sys, os, json, argparse, datetime as dt
from collections import defaultdict, Counter
from openpyxl import load_workbook
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfcommon import SHEETS, BUS, canon_bu, norm_ref, norm_supplier, to_date, num, fmt

ap = argparse.ArgumentParser()
ap.add_argument('master'); ap.add_argument('--template', required=True); ap.add_argument('--out', required=True)
ap.add_argument('--ocp'); ap.add_argument('--sap'); ap.add_argument('--log'); ap.add_argument('--from', dest='frm', default='2026-10-01'); ap.add_argument('--json')
A = ap.parse_args(); FROM = dt.datetime.strptime(A.frm, '%Y-%m-%d')
R = {'ocp': {}, 'sap': {}, 'log': {}, 'needs_ruling': []}
pack = load_workbook(A.template)
def put(tab, rows):
    ws = pack[tab]
    for i, row in enumerate(rows, 4):
        for j, v in enumerate(row, 1): ws.cell(i, j, v)

print('Reading master …'); mw = load_workbook(A.master, read_only=True, data_only=True)

# ---------------- 1. OCP collections
if A.ocp:
    have = {}; 
    for r in mw[SHEETS['cashin']].iter_rows(min_row=2, max_col=42, values_only=True):
        d = to_date(r[29])
        if d and dt.datetime.combine(d, dt.time()) >= FROM: have[(norm_ref(r[8]), round(num(r[25]), 2))] = (r[40], r[41], r[30])
    ow = load_workbook(A.ocp, read_only=True, data_only=True); ws = ow['Payments Key']
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]; ix = {h: i for i, h in enumerate(hdr)}
    rows = []; seen_new = {}; tot = 0.0; by_bu = Counter(); unmapped = 0; n_in_file = 0; file_keys = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = r[ix['Date Paiement']]
        if not isinstance(d, dt.datetime) or d < FROM: continue
        ref, amt = r[ix['Réference']], num(r[ix['Net à payer']]); key = (norm_ref(ref), round(amt, 2)); n_in_file += 1; file_keys.add(key)
        if key in have: continue
        cls = str(r[ix['Class']] or ''); flag = 'Labor' if cls.startswith('Proserv') else ('Non-Labor' if cls.startswith('Procure') else '')
        bu = canon_bu(r[ix['BU']]); 
        if bu not in BUS: unmapped += 1; bu_out = 'Unmapped'
        else: bu_out = bu
        rows.append([ref, r[ix['Fournisseur']], amt, d, r[ix['Réf Paiement']], r[ix['Pièce comptable']], bu_out, flag, r[ix['Division']], None, None, r[ix['ID DOC']],
                     f"OCP Payments Weekly.xlsx / Payments Key (extract {str(r[ix['Extract date']])[:10]}); BU/class per OCP weekly mapping" + ('' if bu in BUS else f" — BU '{r[ix['BU']]}' not a report BU")])
        tot += amt; by_bu[(bu_out, flag)] += amt
    gone = [(k, v) for k, v in have.items() if k not in file_keys]
    put('1_CashIn_OCP', rows)
    R['ocp'] = {'invoices_in_file_from': n_in_file, 'already_in_master': n_in_file - len(rows), 'new_rows': len(rows), 'new_amount': round(tot, 2),
                'new_by_bu': {f'{k[0]} / {k[1]}': round(v, 2) for k, v in by_bu.items()}, 'unmapped_new': unmapped,
                'in_master_not_in_file': [(k[0], k[1], str(v[2])[:10]) for k, v in gone][:20], 'in_master_not_in_file_count': len(gone)}
    if unmapped: R['needs_ruling'].append(f"OCP: {unmapped} new invoices have no report BU (loaded as Unmapped, excluded from BU blocks) — see OCP file 'Unmapped Invoices'")
    if gone: R['needs_ruling'].append(f"OCP: {len(gone)} October invoices in the master are not in the new OCP file (same ref+amount) — check reversals/corrections")
    print(f"OCP: {n_in_file} invoices since {A.frm} in file, {len(rows)} new ({fmt(tot)} MAD), {unmapped} unmapped, {len(gone)} master lines absent from file")

# ---------------- 2. SAP payments
if A.sap:
    sw = load_workbook(A.sap, read_only=True, data_only=True); name = [n for n in sw.sheetnames if n.upper().startswith('PAYMENTS')][0]; ws = sw[name]
    hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]; ix = {}
    for i, h in enumerate(hdr): ix.setdefault(h, i)
    BUMAP = {'WATER': 'Water & Environement', 'WEN': 'Water & Environement', 'F&C': 'Fertilizers & Chemicals', 'FEC': 'Fertilizers & Chemicals', 'AM': 'AM', 'AMT': 'AM', 'ENE': 'Energy Operations',
             'SP2M': 'SP2M Ad Hoc Program', 'GMM': 'Global Mining', 'MINING': 'Global Mining', 'JET': 'JESA Technologies', 'B&I': 'Building & Infrastructure', 'BIN': 'Building & Infrastructure', 'ADS': 'Advisory', 'OVERHEAD': 'Overhead (corporate)'}
    NAT = {'1 - National': '1', '2 - National': '2', '3 - National': '3', '1 - International': 'Intl 1', '2 - International': 'Intl 2', '3 - International': 'Intl 3', 'Agency Etranger': 'Agency Local', 'Filiale OCP': 'Filiale'}
    CODE = {'AMT': 'AM', 'AM': 'AM', 'F&C': 'Fertilizers & Chemicals', 'FEC': 'Fertilizers & Chemicals', 'WATER': 'Water & Environement', 'WEN': 'Water & Environement', 'GMM': 'Global Mining',
            'SP2M': 'SP2M Ad Hoc Program', 'ENERGY': 'Energy Operations', 'ENE': 'Energy Operations', 'ADS': 'Advisory', 'JET': 'JESA Technologies', 'BIN': 'Building & Infrastructure', 'B&I': 'Building & Infrastructure',
            'BIC': 'Building & Capabilities', 'OVERHEAD': 'Overhead (corporate)', 'CORPORATE': 'Overhead (corporate)', 'PTX': 'Power to X (not a report BU)'}
    # invoice (facture) lines grouped by clearing document: reference, amount, BU code
    fac = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[ix['Document Number']] is None or not str(r[ix['Type']] or '').lower().startswith('facture'): continue
        fac[str(r[ix['Clearing Document']] or '').strip()].append((r[ix['Reference']], num(r[ix['Montant en MAD']]), str(r[ix['BU']] or '').strip().upper()))
    # supplier history: BU used for the same supplier in the master's September / October payment tabs
    HIST = defaultdict(Counter)
    for tab, cbu, cname in ((SHEETS['pay_oct'], 33, 6), (SHEETS['pay_sep'], 36, 5)):
        for r in mw[tab].iter_rows(min_row=2, max_col=40, values_only=True):
            b = r[cbu]
            if b in BUS and r[cname]: HIST[norm_supplier(r[cname])][b] += 1
    rows = []; types = Counter(); mx = None; tot_pos = 0.0; bu_src = Counter()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[ix['Document Number']] is None: continue
        t = str(r[ix['Type']] or '').strip(); types[t.lower()] += 1
        if not t.lower().startswith('paiement'): continue
        clr = str(r[ix['Clearing Document']] or '').strip(); lines = fac.get(clr, []) if clr else []
        # net cash per invoice reference = -(sum of all invoice-line amounts with that reference: invoice TTC negative, WHT/retention positive)
        net = defaultdict(float); order = []
        for ref, amt, b in lines:
            if ref in (None, ''): continue
            if ref not in net: order.append(ref)
            net[ref] += -amt
        inv = [(ref, net[ref]) for ref in order if net[ref] > 0]
        refs = ';'.join(str(x[0]) for x in inv); amts = ';'.join(f"{x[1]:.2f}" for x in inv)
        bus = Counter()
        for ref, amt, b in lines:
            if b: bus[CODE.get(b, b)] += abs(amt)
        own = str(r[ix['BU']] or '').strip().upper()
        if own: bus[CODE.get(own, own)] += 1
        note = ''
        if bus:
            top, _ = bus.most_common(1)[0]; bu = top
            if len(bus) > 1:
                note = 'mixed BUs on invoice lines: ' + ', '.join(f"{k} {v:,.0f}" for k, v in bus.most_common()) + f" — dominant {top} used; confirm or split"
                bu_src['invoice lines (mixed, dominant)'] += 1
            else: bu_src['invoice lines'] += 1
        else:
            h = HIST.get(norm_supplier(r[ix['Name']]))
            if h: bu = h.most_common(1)[0][0]; note = f"supplier history ({bu}: {h[bu]} prior lines)"; bu_src['supplier history'] += 1
            else: bu = ''; bu_src['none'] += 1
        if bu not in BUS and bu != 'Overhead (corporate)' and bu: note = (note + '; ' if note else '') + f"'{bu}' is not a report BU"; bu = ''
        cd = r[ix['Clearing Date']]; mx = cd if isinstance(cd, dt.datetime) and (mx is None or cd > mx) else mx
        tier_raw = str(r[0] or '').strip(); tier = NAT.get(tier_raw, tier_raw if tier_raw in ('Not on OCP Tier', 'Office/Publique', 'OCP et Filiale', 'Agency Local', 'Filiale') else ('' if tier_raw == 'International' else tier_raw))
        pos = str(r[ix['Paiement positive']] or '')
        if pos == 'Oui': tot_pos += num(r[ix['Montant en MAD']])
        rows.append([t[0].upper() + t[1:], num(r[ix['Montant en MAD']]), pos, r[ix['Supplier']], r[ix['Name']], r[ix['Document Number']], r[ix['Document Type']], r[ix['Reference']],
                     r[ix['Clearing Document']], cd, r[ix['Posting Date']], r[ix['Purchasing Document']] or r[ix.get('PO', ix['Purchasing Document'])], bu, tier,
                     f"{os.path.basename(A.sap)} / {name}; BU from invoice lines of clearing doc {clr}" + (f"; {note}" if note else '') + (f"; New Tier '{tier_raw}'" if tier_raw == 'International' else ''),
                     refs, amts])
    ws2 = pack['2_CashOut_SAP']; ws2.cell(3, 16, 'Invoice refs (invoice lines of the clearing doc, ;-separated)'); ws2.cell(3, 17, 'Invoice amounts (;-separated, positive)')
    put('2_CashOut_SAP', rows)
    R['sap'] = {'payment_lines': len(rows), 'max_clearing_date': str(mx)[:10], 'types_in_file': dict(types), 'paiement_oui_amount_gross': round(tot_pos, 2), 'bu_source': dict(bu_src)}
    print(f"SAP: {len(rows)} payment lines to {str(mx)[:10]} (invoice lines skipped: {sum(v for k, v in types.items() if k.startswith('facture'))})")

# ---------------- 3b. Tracking log check (per open ongoing line) + log-open lines not on the position
if A.log:
    lw = load_workbook(A.log, read_only=True, data_only=True); ws = lw['Invoice Tracker']
    hdr = None; LOG = []
    for r in ws.iter_rows(min_row=1, max_row=5, values_only=True):
        if r and 'Request date' in [str(x).strip() for x in r if x]: hdr = [str(x).strip() if x else '' for x in r]; break
    ix = {}; 
    for i, h in enumerate(hdr): ix.setdefault(h, i)
    def g(r, *pref):
        for p in pref:
            for h, i in ix.items():
                if h.lower().startswith(p.lower()): return r[i]
        return None
    for r in ws.iter_rows(min_row=4, values_only=True):
        if g(r, 'Invoice Reference') in (None, '') and g(r, 'Supplier Name') in (None, ''): continue
        LOG.append(dict(ref=g(r, 'Invoice Reference'), key=norm_ref(g(r, 'Invoice Reference')), sup=g(r, 'Supplier Name'), bu=g(r, 'BU'),
                        amt=num(g(r, 'AP. MAD Amount') or g(r, 'Ops. MAD Amount') or g(r, 'Invoice Amount')), status=str(g(r, 'Overall status') or '').strip(),
                        ap_block=g(r, 'AP : Type', 'Type de blocage'), treso=g(r, 'Treasury Comments'), tblock=g(r, 'Treausry', 'Treasury : Type'), pdate=g(r, 'Date', 'Payment Date'),
                        req=g(r, 'Request date'), doc=g(r, 'Document Number'), po=g(r, 'PO')))
    by_key = defaultdict(list)
    for l in LOG: by_key[l['key']].append(l)
    og = mw[SHEETS['ongoing']]; hdr_og = {str(c.value).strip(): i for i, c in enumerate(next(og.iter_rows(min_row=1, max_row=1))) if c.value}
    def oc(p):
        for k, i in hdr_og.items():
            if k.lower().startswith(p.lower()): return i
    c_ref, c_sup, c_amt, c_stat, c_in, c_bu, c_key = (oc(x) for x in ['Reference', 'Supplier', 'Amount', 'Sept payment status', 'In payment tracking', 'BU', 'Ref key'])
    rows = []; st = Counter(); matched_keys = set(); newly = 0
    for i, r in enumerate(og.iter_rows(min_row=2, values_only=True), 2):
        if r[0] is None or not str(r[c_stat] or '').startswith('Open'): continue
        key = norm_ref(r[c_ref]); cands = by_key.get(key, [])
        best = None
        for l in cands:
            if norm_supplier(l['sup'])[:8] == norm_supplier(r[c_sup])[:8] or abs(abs(l['amt']) - abs(num(r[c_amt]))) <= 1: best = l; break
        if best is None and cands: best = cands[0]; weak = True
        else: weak = False
        if best:
            matched_keys.add(key); st[best['status'] or '(blank)'] += 1
            if str(r[c_in]) != 'Yes': newly += 1
            rows.append([i, r[c_ref], r[c_sup], 'Yes', best['status'], best['ap_block'], best['treso'], best['tblock'], best['pdate'], best['amt'], 'WEAK MATCH (ref only)' if weak else 'ref + supplier/amount', os.path.basename(A.log)])
        else:
            rows.append([i, r[c_ref], r[c_sup], 'No', None, None, None, None, None, None, 'not in log', os.path.basename(A.log)])
    # extend template tab if missing
    if '3b_LogCheck' not in pack.sheetnames:
        ws3 = pack.create_sheet('3b_LogCheck'); ws3['A1'] = '3b LogCheck — tracking-log status for every open ongoing line (built by build_pack_from_sources.py)'
        for j, h in enumerate(['Ongoing row', 'Reference / Invoice', 'Supplier', 'In log? (Yes/No)', 'Overall status', 'AP blocage', 'Treasury comment', 'Treasury blocage', 'Payment date', 'Log amount (MAD)', 'Match basis', 'Source'], 1): ws3.cell(3, j, h)
    put('3b_LogCheck', rows)
    open_not_on_pos = [l for l in LOG if l['status'].upper() not in ('PAID', 'CANCELLED', 'CANCELED') and l['key'] not in matched_keys and l['key']]
    # also exclude those whose key exists on the position in a closed state (paid etc.)
    pos_keys = {norm_ref(r[c_ref]) for r in og.iter_rows(min_row=2, values_only=True) if r[0] is not None}
    open_not_on_pos_new = [l for l in open_not_on_pos if l['key'] not in pos_keys]
    agg = defaultdict(lambda: [0, 0.0])
    for l in open_not_on_pos_new: agg[(canon_bu(l['bu']) or str(l['bu']), l['status'])][0] += 1; agg[(canon_bu(l['bu']) or str(l['bu']), l['status'])][1] += l['amt']
    R['log'] = {'log_lines': len(LOG), 'open_ongoing_lines_checked': len(rows), 'found_in_log': sum(1 for x in rows if x[3] == 'Yes'), 'newly_flagged_in_log': newly,
                'status_of_matched': dict(st), 'log_open_not_on_position': {'lines': len(open_not_on_pos_new), 'amount': round(sum(l['amt'] for l in open_not_on_pos_new), 2),
                'by_bu_status': {f'{k[0]} | {k[1]}': [v[0], round(v[1], 2)] for k, v in sorted(agg.items())}},
                'log_open_lines_sample': [(l['bu'], l['sup'], l['ref'], round(l['amt'], 2), l['status'], str(l['req'])[:10]) for l in open_not_on_pos_new[:40]]}
    if open_not_on_pos_new: R['needs_ruling'].append(f"Tracking log: {len(open_not_on_pos_new)} unpaid log lines ({fmt(sum(l['amt'] for l in open_not_on_pos_new))} MAD) are not on the ongoing position — add as approved? (listed in the report)")
    print(f"LOG: {len(LOG)} lines; {len(rows)} open ongoing lines checked, {R['log']['found_in_log']} in log ({newly} newly flagged); {len(open_not_on_pos_new)} unpaid log lines not on the position")

pack.save(A.out); print('pack written', A.out)
if A.json: json.dump(R, open(A.json, 'w'), indent=1, default=str, ensure_ascii=False)
