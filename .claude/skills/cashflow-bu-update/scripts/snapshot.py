#!/usr/bin/env python3
"""Print the state of a CASH FLOW ANALYSIS - BUs workbook: controls, headline, integrity checks,
as-of per input stream, last Data Sources rows. Reads cached values (last Excel save).
Usage: snapshot.py <workbook.xlsx> [--json out.json]"""
import sys, json, os, datetime as dt
from collections import Counter
from openpyxl import load_workbook
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfcommon import SHEETS, to_date, fmt

path = sys.argv[1]
wb = load_workbook(path, read_only=True, data_only=True)
out = {'file': os.path.basename(path), 'size_mb': round(os.path.getsize(path) / 1e6, 1)}

ex = wb[SHEETS['exec']]
out['controls'] = {'include_ongoing': ex['F6'].value, 'data_as_of': str(ex['E7'].value)[:10], 'basis': ex['F8'].value,
                   'title_line': ex['C3'].value}
labels = ['collections_ocp', 'dividends', 'cash_out_jul_sep', 'cash_out_oct_actual', 'approved_pmts', 'net_after_approved',
          'pending_pmts', 'net_all_ongoing']
out['headline'] = {k: ex.cell(13, 3 + i).value for i, k in enumerate(labels)}
out['psr'] = {'psr_collections': ex['C16'].value, 'psr_reserve': ex['D16'].value, 'am_psr_excluded': ex['E16'].value,
              'net_psr': ex['G16'].value}
out['by_month'] = {}
for r, k in [(24, 'total_cash_in'), (26, 'suppliers'), (27, 'lc'), (28, 'customs'), (29, 'exclusions_added_back'),
             (30, 'total_cash_out'), (31, 'net_position')]:
    out['by_month'][k] = {m: ex.cell(r, c).value for m, c in zip(['Jul', 'Aug', 'Sep', 'Oct_actual', 'Oct_ongoing', 'Cumulative'], range(4, 10))}

ck = wb[SHEETS['checks']]
out['checks'] = [{'check': ck.cell(r, 3).value, 'diff': ck.cell(r, 4).value, 'status': ck.cell(r, 6).value}
                 for r in range(23, 41) if ck.cell(r, 3).value]
out['checks_not_ok'] = [c['check'] for c in out['checks'] if str(c['status']).strip().upper() != 'OK']

def col_stats(sheet, date_col, amt_col=None, flt=None, maxc=60):
    ws = wb[sheet]; n = 0; mx = None; tot = 0.0
    for row in ws.iter_rows(min_row=2, max_col=maxc, values_only=True):
        if flt and not flt(row): continue
        d = to_date(row[date_col]) if date_col is not None else None
        if date_col is not None and d is None: continue
        n += 1
        if d and (mx is None or d > mx): mx = d
        if amt_col is not None and isinstance(row[amt_col], (int, float)): tot += row[amt_col]
    return {'rows': n, 'max_date': str(mx) if mx else None, 'amount': round(tot, 2)}

streams = {}
streams['1_cash_in_ocp'] = col_stats(SHEETS['cashin'], 29, 25, maxc=46)           # AD date, Z amount
streams['1_cash_in_ocp_october'] = col_stats(SHEETS['cashin'], 29, 25, flt=lambda r: (to_date(r[29]) or dt.date(1900,1,1)) >= dt.date(2026,10,1), maxc=46)
streams['2_sap_payments_oct_included'] = col_stats(SHEETS['pay_oct'], 17, 2, flt=lambda r: r[31] == 'Yes', maxc=37)  # R clearing date, C amount, AF include
streams['2_sap_payments_oct_all'] = col_stats(SHEETS['pay_oct'], 17, 2, maxc=37)
streams['5a_lc_paid'] = col_stats(SHEETS['lc'], 4, 6, maxc=37)                      # E date, G debit
streams['6_customs_dd'] = col_stats(SHEETS['customs'], 1, 4, maxc=13)               # B date, E amount
og = wb[SHEETS['ongoing']]
st = Counter(); appr = Counter(); amt = Counter(); last_added = None
for row in og.iter_rows(min_row=2, max_col=34, values_only=True):
    if row[0] is None: continue
    s = str(row[13] or ''); a = str(row[26] or '')
    key = 'open' if s.startswith('Open') else 'closed'
    st[key] += 1
    if key == 'open':
        appr['approved' if a == 'Yes' else 'pending'] += 1
        amt['approved' if a == 'Yes' else 'pending'] += float(row[4] or 0)
    d = to_date(row[24])
    if d and d < dt.date(2027, 1, 1) and (last_added is None or d > last_added): last_added = d
streams['3_4_ongoing_position'] = {'open_lines': st['open'], 'closed_lines': st['closed'], 'approved_lines': appr['approved'],
                                   'approved_amount': round(amt['approved'], 2), 'pending_lines': appr['pending'],
                                   'pending_amount': round(amt['pending'], 2), 'last_date_added': str(last_added)}
lo = wb[SHEETS['lc_ongoing']]
streams['5b_lc_situation'] = {'tab': SHEETS['lc_ongoing'], 'title': lo['A1'].value, 'update_note': lo['A3'].value}
out['streams'] = streams

ds = wb[SHEETS['sources']]
rows = [r for r in ds.iter_rows(min_row=2, max_col=11, values_only=True) if r[0]]
out['data_sources_rows'] = len(rows)
out['data_sources_last'] = [{'written_to': str(r[0])[:70], 'source': str(r[1])[:70], 'run_at': str(r[6])[:20], 'rows': str(r[8])[:30]} for r in rows[-6:]]

if '--json' in sys.argv:
    json.dump(out, open(sys.argv[sys.argv.index('--json') + 1], 'w'), indent=1, default=str)

print(f"FILE {out['file']} ({out['size_mb']} MB)")
print(f"Data as of {out['controls']['data_as_of']} · include ongoing = {out['controls']['include_ongoing']}")
print(f"Basis: {out['controls']['basis']}")
print("\nHEADLINE (MAD)")
for k, v in out['headline'].items(): print(f"  {k:22s} {fmt(v):>20s}")
print("\nINTEGRITY CHECKS:", 'ALL OK' if not out['checks_not_ok'] else 'NOT OK -> ' + '; '.join(out['checks_not_ok']))
print("\nSTREAMS")
for k, v in streams.items(): print(f"  {k:32s} {v}")
print(f"\nDATA SOURCES: {out['data_sources_rows']} rows; last:")
for r in out['data_sources_last']: print(f"  {r['run_at']:20s} {r['written_to']:70s} <- {r['source']}")
