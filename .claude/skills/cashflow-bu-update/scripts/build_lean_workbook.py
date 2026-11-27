#!/usr/bin/env python3
"""Build the lean 'COO Cash Position' workbook from the full analysis workbook:
Summary · Cash IN (flat table) · Cash OUT (SAP layout, one table for all months incl. LC and customs) · Ongoing (live columns only)
· Tiering Master · OH Suppliers · Log · README.  All data tabs are values; only the Summary and the Month columns hold formulas.
Usage: build_lean_workbook.py <full.xlsx> <lean.xlsx> [--as-of 2026-11-02]"""
import sys, os, re, datetime as dt, json
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter as L
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cfcommon import BUS, BU_CODES, LC_PROJECT, norm_supplier, to_date, num, tier_str

SRC, OUT = sys.argv[1], sys.argv[2]
AS_OF = sys.argv[sys.argv.index('--as-of') + 1] if '--as-of' in sys.argv else dt.date.today().isoformat()
print('reading', SRC); wb = load_workbook(SRC, read_only=True, data_only=True)
F = 'Arial'; HDR_FILL = PatternFill('solid', start_color='1F3864'); HDR_FONT = Font(name=F, bold=True, color='FFFFFF', size=10)
SUB_FILL = PatternFill('solid', start_color='D9E1F2'); INPUT_FONT = Font(name=F, color='0000FF', size=10); BOLD = Font(name=F, bold=True, size=10)
NUM = '#,##0;(#,##0);-'; NUM2 = '#,##0.00;(#,##0.00);-'; DATE = 'dd/mm/yyyy'
thin = Side(style='thin', color='BFBFBF')
OH = set()
for r in wb['OH Suppliers'].iter_rows(min_row=2, values_only=True):
    if len(r) > 4 and r[4]: OH.add(norm_supplier(r[4]))
def is_oh(name): return norm_supplier(name) in OH
def month(d): d = to_date(d); return dt.datetime(d.year, d.month, 1) if d else None

# ------------------------------------------------------------------ CASH IN
cash_in = []
for r in wb['SRC - Cash In (AX Jul-Aug)'].iter_rows(min_row=2, max_col=46, values_only=True):
    d = to_date(r[29]); amt = r[25]
    if d is None or d < dt.date(2026, 7, 1) or not isinstance(amt, (int, float)): continue
    bu = r[40] or 'Unmapped'; flag = r[41]
    cls = 'PSR' if flag == 'Labor' else ('PROC' if flag == 'Non-Labor' else ('PROC (unclassified)' if bu != 'Unmapped' else 'Unmapped'))
    cash_in.append([dt.datetime.combine(d, dt.time()), r[8], r[6], r[2], r[28], r[30], r[20], float(amt), cls, bu, 'OCP payment report (AX extract)' + (f' · {str(r[45])[:60]}' if r[45] else '')])
pc = wb['PROC Compensable (standing)']
for rr in range(5, 15):
    b, v = pc.cell(rr, 1).value, pc.cell(rr, 3).value
    if isinstance(v, (int, float)) and v: cash_in.append([dt.datetime(2026, 9, 30), f'Dividend — {b}', 'PROC compensable (standing)', None, None, None, None, float(v), 'Dividend', b, 'Standing input 17/09/2026 (PROC compensable by BU)'])
cash_in.sort(key=lambda x: (x[0], str(x[9])))
print('cash in rows', len(cash_in))

# ------------------------------------------------------------------ CASH OUT (SAP layout)
SAP_COLS = ['New Tier', 'Type', 'Montant en MAD', 'Paiement positive', 'Company Code', 'Supplier', 'Name', 'Document Number', 'Document Type', 'Document Date', 'Special G/L Ind.',
            'Company Code Currency Value', 'Company Code Currency Key', 'Document Currency Value', 'Document Currency Key', 'Reference', 'Clearing Document', 'Clearing Date', 'User Name',
            'G/L Account', 'Text', 'Offsetting Account', 'Posting Date', 'Entry Date', 'Reference Key', 'Reference Key 1', 'Reference Key 2', 'Reference Key 3', 'Purchasing Document']
HELP = ['Month', 'Kind', 'Include?', 'Cash out (MAD)', 'Tier', 'BU', 'OH?', 'Source / note']
cash_out = []
def row_sap(vals, month_, kind, include, tier, bu, oh, note):
    base = [vals.get(c) for c in SAP_COLS]
    return base + [month_, kind, include, None, tier, bu, oh, note]
# July / August (mapped extracts: BU, Tiering, Nature, Company Code, Supplier Code, Supplier Name, Document Number, Document Type, Posting Date, Clearing Date, Year/Month, Amount in MAD, Line Text, OPEX/CAPEX, OH?, WBS)
for tab, m in (('SRC - Payments Jul (Mapped)', dt.datetime(2026, 7, 1)), ('SRC - Payments Aug (Mapped)', dt.datetime(2026, 8, 1))):
    for r in wb[tab].iter_rows(min_row=2, max_col=16, values_only=True):
        if not isinstance(r[11], (int, float)): continue
        bu = r[0] or ''; tier = tier_str(r[1]); oh = 'OH' if (r[14] == 'OH' or is_oh(r[5])) else ''
        inc = 'Yes' if bu else 'No - no BU'
        if inc == 'Yes' and tier not in ('1','2','3','Intl 1','Intl 2','Intl 3','Etranger*','OCP et Filiale','Office/Publique','Not on OCP Tier','Unclassified'): inc = 'No - excluded tier'
        if bu in ('AM', 'Water & Environement') and oh == 'OH': inc = 'No - overhead excluded (AM/Water)'
        vals = {'New Tier': r[1], 'Type': r[2], 'Montant en MAD': -float(r[11]), 'Company Code': r[3], 'Supplier': r[4], 'Name': r[5], 'Document Number': r[6], 'Document Type': r[7],
                'Posting Date': r[8], 'Clearing Date': r[9], 'Text': r[12]}
        cash_out.append(row_sap(vals, m, 'Supplier payment (SAP)', inc, tier, bu or '(unmapped)', oh, f'{tab} (Payments July-Sept.8 FINAL)'))
# September (SAP layout A:AE + helpers AK BU, AM tier, AP OH)
ws = wb['September Actual Payments']; hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
for r in ws.iter_rows(min_row=2, max_col=47, values_only=True):
    if not isinstance(r[2], (int, float)) or r[2] == 0: continue
    bu = r[36] or '(unmapped)'; tier = tier_str(r[38]); oh = 'OH' if r[41] == 'OH' else ''
    inc = 'Yes' if bu != '(unmapped)' else 'No - no BU'
    if inc == 'Yes' and tier not in ('1','2','3','Intl 1','Intl 2','Intl 3','Etranger*','OCP et Filiale','Office/Publique','Not on OCP Tier','Unclassified'): inc = 'No - excluded tier'
    if bu in ('AM', 'Water & Environement') and oh == 'OH': inc = 'No - overhead excluded (AM/Water)'
    vals = {c: r[i] for i, c in enumerate(hdr[:31]) if c in SAP_COLS}; vals['Montant en MAD'] = -float(r[2])
    if vals.get('Name') in (None, '') and r[5] not in (None, ''): vals['Name'] = r[5]
    if vals.get('Reference') in (None, '') and r[14] not in (None, ''): vals['Reference'] = r[14]
    cash_out.append(row_sap(vals, dt.datetime(2026, 9, 1), 'Supplier payment (SAP)', inc, tier, bu, oh, 'September Actual Payments (SAP situation 28/09 + 02/10 loads)' + ('' if r[7] else ' · line without SAP document (manual/tracker line)')))
# October (SAP layout + Include flag)
ws = wb['October Actual Payments']; hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]; ix = {h: i for i, h in enumerate(hdr)}
for r in ws.iter_rows(min_row=2, max_col=37, values_only=True):
    if r[7] is None or not isinstance(r[2], (int, float)): continue
    inc = str(r[31] or 'No'); bu = r[33] or '(unmapped)'; tier = tier_str(r[32]); oh = 'OH' if r[35] == 'OH' else ''; name = str(r[6] or '')
    if inc == 'Yes' and bu != '(unmapped)' and tier not in ('1','2','3','Intl 1','Intl 2','Intl 3','Etranger*','OCP et Filiale','Office/Publique','Not on OCP Tier','Unclassified'): inc = 'No - excluded tier'
    if inc == 'Yes' and bu in ('AM', 'Water & Environement') and oh == 'OH': inc = 'No - overhead excluded (AM/Water)'
    if inc == 'Yes' and bu == 'AM' and name.upper().startswith('PISTELLI'): inc = 'No - AM PISTELLI (cash-in recovery)'
    vals = {c: r[i] for i, c in enumerate(hdr[:30]) if c in SAP_COLS}; vals['Montant en MAD'] = float(r[2])
    cash_out.append(row_sap(vals, month(r[17]) or dt.datetime(2026, 10, 1), 'Supplier payment (SAP)', inc, tier, bu, oh, f"October Actual Payments · {str(r[34] or '')[:80]}"))
# LC payments (bank statement)
for r in wb['SRC - LC Payments'].iter_rows(min_row=2, max_col=37, values_only=True):
    d = to_date(r[4]); amt = r[6]
    if d is None or not isinstance(amt, (int, float)): continue
    proj = str(r[18] or '').strip(); bu = BU_CODES.get(proj) if proj in BU_CODES else None
    inc = 'Yes' if bu and proj in LC_PROJECT.values() else 'No - LC without report BU'
    if proj == 'AMT' and str(r[23] or '').upper().startswith('PELZ') and d >= dt.date(2026, 10, 1): inc = 'No - AM PELZ (cash-in recovery)'
    vals = {'Type': 'LC payment', 'Montant en MAD': float(amt), 'Company Code': r[1], 'Name': r[13], 'Document Number': r[12], 'Reference': r[12], 'Clearing Date': dt.datetime.combine(d, dt.time()),
            'Text': r[5], 'Document Currency Value': r[8], 'Document Currency Key': r[9], 'Purchasing Document': r[10]}
    cash_out.append(row_sap(vals, month(d), 'LC payment (bank)', inc, 'LC', bu or f'({proj})', '', 'SRC - LC Payments (bank statement LC lines)'))
# Customs (Treasury detail)
for r in wb['DD'].iter_rows(min_row=2, max_col=13, values_only=True):
    d = to_date(r[1]); amt = r[4]
    if d is None or not isinstance(amt, (int, float)): continue
    bu = r[8] or BU_CODES.get(str(r[7] or '').upper(), str(r[7]))
    vals = {'Type': 'Customs', 'Montant en MAD': float(amt), 'Company Code': r[0], 'Name': 'RECEVEUR DOUANE (TGR)', 'Document Number': r[3], 'Text': r[3], 'Clearing Date': dt.datetime.combine(d, dt.time()), 'Document Date': r[2]}
    cash_out.append(row_sap(vals, month(d), 'Customs (Treasury)', 'Yes' if bu in BUS else 'No - no BU', 'Customs', bu, '', 'DD (Treasury customs detail by BU)'))
cash_out.sort(key=lambda x: (x[29] or dt.datetime(2099, 1, 1), x[30], str(x[17] or '')))
print('cash out rows', len(cash_out))

# ------------------------------------------------------------------ ONGOING (live columns only)
og = wb['REPORT - Oct Ongoing-Position']; hdr = [str(c.value).strip() if c.value else '' for c in next(og.iter_rows(min_row=1, max_row=1))]
def col(prefix):
    for i, h in enumerate(hdr):
        if h.lower().startswith(prefix.lower()): return i
    return None
cBU, cSup, cPO, cRef, cAmt, cDue, cAppr, cTier, cPrio, cCom, cSrc, cStat, cOvr, cLog, cLogSt, cAdded, cOld = (col(x) for x in ['BU', 'Supplier', 'PO', 'Reference', 'Amount', 'Due date', 'Approved? (Yes', 'Tier', 'Priority', 'Comment', 'Source', 'Sept payment status', 'Manual status override', 'In payment tracking log', 'Tracking log — Overall', 'Date added', 'Approved payments (Oct)'])
def simple_status(s):
    s = str(s or '').strip()
    if s.startswith('Open'): return 'Open'
    if s.startswith('Paid in Sept'): return 'Paid in Sept'
    if s.startswith('Paid in Oct'): return 'Paid in Oct'
    if s.startswith('Paid'): return 'Paid'
    for k in ('Removed', 'Superseded', 'Deferred', 'Excluded', 'Duplicate'):
        if s.startswith(k): return k
    return s or 'Open'
ongoing = []
for r in og.iter_rows(min_row=2, max_col=len(hdr), values_only=True):
    if r[cBU] is None and r[cSup] is None: continue
    raw = r[cStat] if r[cStat] not in (None, '') else r[cOvr]; st = simple_status(raw)
    appr = r[cAppr] if cAppr is not None and r[cAppr] in ('Yes', 'No') else (r[cOld] if cOld is not None else None)
    if st != 'Open': appr = ''
    ongoing.append([r[cBU], r[cSup], r[cPO], r[cRef], num(r[cAmt]) if r[cAmt] not in (None, '') else None, r[cDue], tier_str(r[cTier]), st, appr, r[cLog], r[cPrio], r[cAdded], (r[cSrc] or '')[:200] if isinstance(r[cSrc], str) else r[cSrc], str(raw or '')[:120], (r[cCom] or '')[:250] if isinstance(r[cCom], str) else r[cCom]])
# LC maturities (approved ongoing)
lo = wb['LC Ongoing (19-10)']
for rr in range(5, lo.max_row + 1):
    if str(lo.cell(rr, 1).value or '').upper().startswith('TOTAL'): break
    if lo.cell(rr, 3).value in (None, ''): continue
    q = lo.cell(rr, 17).value; bu = lo.cell(rr, 16).value
    if not isinstance(q, (int, float)) or not bu or q == 0: continue
    ongoing.append([bu, lo.cell(rr, 6).value, lo.cell(rr, 4).value, lo.cell(rr, 5).value, float(q), None, 'LC', 'Open', 'Yes', 'n/a', 'LC maturity (bank situation 19/10)', dt.datetime(2026, 10, 20), 'LCs SITUATION - October 19th (003).xlsx', 'Open - LC maturity', lo.cell(rr, 14).value])
order = {('Open', 'Yes'): 0, ('Open', 'No'): 1}
ongoing.sort(key=lambda x: (order.get((x[7], x[8]), 2), str(x[0]), str(x[1])))
print('ongoing rows', len(ongoing), 'open', sum(1 for x in ongoing if x[7] == 'Open'))

# ------------------------------------------------------------------ reconciliation (python) vs the full workbook's cached headline
ex = wb['EXEC Summary']; head = [ex.cell(13, c).value for c in range(3, 11)]
ci = defaultdict(float)
for r in cash_in: ci[(r[0].strftime('%Y-%m'), r[8])] += r[7]
co = defaultdict(float)
for r in cash_out:
    if r[31] == 'Yes': co[(r[29].strftime('%Y-%m') if r[29] else '?', r[30])] += -r[2]
proc_total = sum(v for (m, c), v in ci.items() if c.startswith('PROC') and m <= '2026-10'); div = sum(v for (m, c), v in ci.items() if c == 'Dividend')
out_js = sum(v for (m, k), v in co.items() if m <= '2026-09'); out_oct = sum(v for (m, k), v in co.items() if m == '2026-10')
appr = sum(x[4] for x in ongoing if x[7] == 'Open' and x[8] == 'Yes'); pend = sum(x[4] for x in ongoing if x[7] == 'Open' and x[8] == 'No')
print('lean by month/kind:', {f'{m} {k}': round(v, 2) for (m, k), v in sorted(co.items())})
rec = {'collections_proc': (round(proc_total, 2), head[0]), 'dividends': (round(div, 2), head[1]), 'cash_out_jul_sep': (round(out_js, 2), head[2]), 'cash_out_oct': (round(out_oct, 2), head[3]), 'approved': (round(appr, 2), head[4]), 'pending': (round(pend, 2), head[6])}
print('RECONCILIATION lean vs full (cached):'); [print(f'  {k:18s} lean {v[0]:>18,.0f}  full {v[1]:>18,.0f}  diff {v[0]-v[1]:>14,.0f}') for k, v in rec.items()]
json.dump({k: {'lean': v[0], 'full': v[1], 'diff': round(v[0] - v[1], 2)} for k, v in rec.items()}, open(os.path.splitext(OUT)[0] + '.reconciliation.json', 'w'), indent=1)

# ------------------------------------------------------------------ write the lean workbook
nb = Workbook(); nb.remove(nb.active)
def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row, c); cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[row].height = 32
def add_table(ws, name, first_row, last_row, ncol):
    t = Table(displayName=name, ref=f"A{first_row}:{L(ncol)}{last_row}"); t.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True); ws.add_table(t)

# README
ws = nb.create_sheet('README'); ws.column_dimensions['A'].width = 4; ws.column_dimensions['B'].width = 125
lines = [('COO CASH POSITION — lean workbook', True), (f'Built {AS_OF} from the full analysis workbook. Data as of: collections to 02/11/2026 (OCP payment report), payments to 30/10/2026 (SAP), ongoing position 03/11/2026.', False), ('', False),
 ('TABS', True), ('Summary — cash IN, cash OUT and ongoing commitments by month. Formulas only (SUMIFS over the three tables). Nothing to type except the data-as-of date.', False),
 ('Cash IN — one line per collected invoice (OCP payment report), Jul 2026 onwards. Class = PSR / PROC / Dividend. Month is a formula.', False),
 ('Cash OUT — one line per SAP payment line in the SAP extract layout (columns A:AC exactly as SAP), for all months, plus LC payments (bank) and customs (Treasury) with Kind = LC / Customs. Helper block at the right: Month (formula), Kind, Include?, Cash out (MAD) (formula), Tier, BU, OH?, Source.', False),
 ('Ongoing — open and closed commitments with the live columns only: BU, Supplier, PO, Reference, Amount, Due date, Tier, Status, Approved?, In tracking log?, Priority/batch, Date added, Source, Status detail, Comment. Open lines first (approved, then pending). LC maturities are included as approved lines.', False),
 ('Tiering Master / OH Suppliers — reference lists. Log — audit trail of every load (from the full workbook).', False), ('', False),
 ('HOW TO UPDATE', True), ('1. Cash IN: paste the new invoices from the OCP payment report at the bottom of the Cash IN table (one row per invoice; fill Class and BU). The Month column fills itself.', False),
 ('2. Cash OUT: paste the new SAP payment lines (paiement lines) at the bottom of the Cash OUT table in the SAP layout (A:AC), then fill Kind = Supplier payment (SAP), Include? = Yes (or No - … for corporate / excluded), Tier, BU. LC payments and customs: one row each with Kind = LC payment (bank) / Customs (Treasury) and the amount in Montant en MAD. Cash out (MAD) and Month are formulas.', False),
 ('3. Ongoing: new commitments = new rows (Status Open, Approved? Yes/No). A payment closes a line by setting Status = Paid in <month>. Never delete a line; use Status = Removed with the reason in Comment.', False),
 ('4. Summary recalculates. Set the data-as-of date in Summary!C4.', False), ('', False),
 ('CONVENTIONS', True), ('Amounts in MAD. Cash IN positive. In Cash OUT, Montant en MAD is as SAP shows it (positive = paid); Cash out (MAD) is the negative cash flow of included lines. Ongoing amounts negative (commitments), credit notes positive.', False),
 ('Include? = Yes counts in the Summary. "No - overhead excluded (AM/Water)" and PISTELLI/PELZ lines follow the standing exclusions of the full workbook; "No - invoice line" are SAP invoice lines kept for reference; "No - corporate OH" are corporate charges.', False),
 ('BU figures: filter any table by BU, or ask for a BU pivot.', False)]
for i, (t, b) in enumerate(lines, 2):
    c = ws.cell(i, 2, t); c.font = Font(name=F, bold=b, size=12 if i == 2 else 10); c.alignment = Alignment(wrap_text=True, vertical='top')

# Cash IN
ws = nb.create_sheet('Cash IN'); hdr_in = ['Date Paiement', 'Invoice reference', 'Fournisseur', 'ID DOC', 'Pièce comptable', 'Réf Paiement', 'Division', 'Net à payer (MAD)', 'Class', 'BU', 'Source / note', 'Month']
ws.append(hdr_in); style_header(ws, 1, len(hdr_in))
for i, r in enumerate(cash_in, 2):
    for j, v in enumerate(r, 1): ws.cell(i, j, v)
    ws.cell(i, 12, f'=DATE(YEAR(A{i}),MONTH(A{i}),1)')
last = len(cash_in) + 1; add_table(ws, 'CashIn', 1, last, 12)
for i in range(2, last + 1): ws.cell(i, 1).number_format = DATE; ws.cell(i, 8).number_format = NUM2; ws.cell(i, 12).number_format = 'mmm-yy'
for c, w in zip('ABCDEFGHIJKL', [12, 18, 30, 10, 14, 14, 16, 16, 18, 24, 50, 9]): ws.column_dimensions[c].width = w
ws.freeze_panes = 'A2'

# Cash OUT
ws = nb.create_sheet('Cash OUT'); hdr_out = SAP_COLS + HELP; ws.append(hdr_out); style_header(ws, 1, len(hdr_out))
for c in range(30, 38): ws.cell(1, c).fill = PatternFill('solid', start_color='7F6000')
for i, r in enumerate(cash_out, 2):
    for j, v in enumerate(r, 1):
        if j == 33: continue
        ws.cell(i, j, v)
    ws.cell(i, 30, f'=IF(R{i}="","",DATE(YEAR(R{i}),MONTH(R{i}),1))'); ws.cell(i, 33, f'=IF(AF{i}="Yes",-C{i},0)')
last = len(cash_out) + 1; add_table(ws, 'CashOut', 1, last, len(hdr_out))
for i in range(2, last + 1):
    for c in (3, 12, 14, 33): ws.cell(i, c).number_format = NUM2
    for c in (10, 18, 23, 24): ws.cell(i, c).number_format = DATE
    ws.cell(i, 30).number_format = 'mmm-yy'
widths = {1: 14, 2: 12, 3: 16, 6: 10, 7: 28, 8: 13, 16: 18, 17: 13, 18: 12, 21: 24, 23: 12, 29: 14, 30: 9, 31: 20, 32: 26, 33: 16, 34: 10, 35: 24, 36: 6, 37: 50}
for c, w in widths.items(): ws.column_dimensions[L(c)].width = w
ws.freeze_panes = 'H2'

# Ongoing
ws = nb.create_sheet('Ongoing'); hdr_og = ['BU', 'Supplier', 'PO', 'Reference / Invoice', 'Amount (MAD)', 'Due date', 'Tier', 'Status', 'Approved? (Yes/No)', 'In tracking log?', 'Priority / batch', 'Date added', 'Source', 'Status detail', 'Comment']
ws.append(hdr_og); style_header(ws, 1, len(hdr_og))
for i, r in enumerate(ongoing, 2):
    for j, v in enumerate(r, 1): ws.cell(i, j, v)
last = len(ongoing) + 1; add_table(ws, 'Ongoing', 1, last, len(hdr_og))
for i in range(2, last + 1): ws.cell(i, 5).number_format = NUM2; ws.cell(i, 6).number_format = DATE; ws.cell(i, 12).number_format = DATE
for c, w in zip(hdr_og, [24, 30, 14, 22, 16, 11, 9, 12, 10, 9, 22, 11, 40, 34, 50]): ws.column_dimensions[L(hdr_og.index(c) + 1)].width = w
ws.freeze_panes = 'E2'

# Summary
ws = nb.create_sheet('Summary', 0)
ws['B2'] = 'COO CASH POSITION — JESA Group'; ws['B2'].font = Font(name=F, bold=True, size=14)
ws['B4'] = 'Data as of'; ws['C4'] = dt.datetime.strptime(AS_OF, '%Y-%m-%d'); ws['C4'].font = INPUT_FONT; ws['C4'].number_format = DATE
ws['D4'] = 'collections per OCP payment report · payments per SAP · ongoing = open commitments (approved / pending approval) · MAD'; ws['D4'].font = Font(name=F, italic=True, size=9, color='595959')
months = [dt.datetime(2026, 7, 1), dt.datetime(2026, 8, 1), dt.datetime(2026, 9, 1), dt.datetime(2026, 10, 1), dt.datetime(2026, 11, 1)]
hdrs = ['MAD', 'Jul-26', 'Aug-26', 'Sep-26', 'Oct-26 (closed)', 'Nov-26 actuals (to date)', 'Nov ongoing — approved', 'Nov ongoing — pending approval', 'Total actuals + ongoing']
for j, h in enumerate(hdrs, 2): ws.cell(6, j, h)
style_header(ws, 6, 10); ws.cell(6, 1).fill = PatternFill(); 
for j, m in enumerate(months, 3): ws.cell(5, j, m).number_format = 'yyyy-mm-dd'; ws.cell(5, j).font = Font(name=F, size=8, color='A6A6A6')
ws.row_dimensions[5].height = 10
def line(row, label, in_formula=None, out_formula=None, bold=False, kind='in'):
    ws.cell(row, 2, label).font = BOLD if bold else Font(name=F, size=10)
    for j in range(3, 8):
        col = L(j)
        f = (in_formula or out_formula)
        if f: ws.cell(row, j, f.replace('{M}', f'{col}$5')).number_format = NUM
r = 7
ws.cell(r, 2, 'CASH IN').font = BOLD; ws.cell(r, 2).fill = SUB_FILL; r += 1
R = {}
R['psr'] = r; line(r, 'PSR (Labor) collections', '=SUMIFS(CashIn[Net à payer (MAD)],CashIn[Month],{M},CashIn[Class],"PSR")'); r += 1
R['proc'] = r; line(r, 'Procurement (Non-Labor) collections', '=SUMIFS(CashIn[Net à payer (MAD)],CashIn[Month],{M},CashIn[Class],"PROC*")'); r += 1
R['div'] = r; line(r, 'Dividends (PROC compensable, standing)', '=SUMIFS(CashIn[Net à payer (MAD)],CashIn[Month],{M},CashIn[Class],"Dividend")'); r += 1
R['in'] = r; line(r, 'Total cash IN', None, None, True)
for j in range(3, 8): ws.cell(r, j, f'=SUM({L(j)}{R["psr"]}:{L(j)}{R["div"]})').number_format = NUM; ws.cell(r, j).font = BOLD
r += 2
ws.cell(r, 2, 'CASH OUT (included lines only)').font = BOLD; ws.cell(r, 2).fill = SUB_FILL; r += 1
R['sup'] = r; line(r, 'Supplier payments (SAP)', '=SUMIFS(CashOut[Cash out (MAD)],CashOut[Month],{M},CashOut[Kind],"Supplier payment (SAP)")'); r += 1
R['lc'] = r; line(r, 'LC payments (bank)', '=SUMIFS(CashOut[Cash out (MAD)],CashOut[Month],{M},CashOut[Kind],"LC payment (bank)")'); r += 1
R['cus'] = r; line(r, 'Customs (Treasury)', '=SUMIFS(CashOut[Cash out (MAD)],CashOut[Month],{M},CashOut[Kind],"Customs (Treasury)")'); r += 1
R['out'] = r; line(r, 'Total cash OUT', None, None, True)
for j in range(3, 8): ws.cell(r, j, f'=SUM({L(j)}{R["sup"]}:{L(j)}{R["cus"]})').number_format = NUM; ws.cell(r, j).font = BOLD
r += 1
R['oh'] = r; line(r, '   memo — overhead / corporate / excluded lines not counted above', '=-SUMIFS(CashOut[Montant en MAD],CashOut[Month],{M},CashOut[Include?],"No - overhead*")-SUMIFS(CashOut[Montant en MAD],CashOut[Month],{M},CashOut[Include?],"No - corporate*")-SUMIFS(CashOut[Montant en MAD],CashOut[Month],{M},CashOut[Include?],"No - AM P*")')
for j in range(3, 8): ws.cell(r, j).font = Font(name=F, italic=True, size=9, color='595959')
r += 2
ws.cell(r, 2, 'ONGOING COMMITMENTS (open lines)').font = BOLD; ws.cell(r, 2).fill = SUB_FILL; r += 1
R['appr'] = r; ws.cell(r, 2, 'Approved payments (in tracking log / approved lists) incl. LC maturities'); ws.cell(r, 8, '=SUMIFS(Ongoing[Amount (MAD)],Ongoing[Status],"Open",Ongoing[Approved? (Yes/No)],"Yes")').number_format = NUM; r += 1
R['pend'] = r; ws.cell(r, 2, 'Pending payment approval'); ws.cell(r, 9, '=SUMIFS(Ongoing[Amount (MAD)],Ongoing[Status],"Open",Ongoing[Approved? (Yes/No)],"No")').number_format = NUM; r += 2
ws.cell(r, 2, 'NET').font = BOLD; ws.cell(r, 2).fill = SUB_FILL; r += 1
R['net'] = r; ws.cell(r, 2, 'Net cash position (IN + OUT)').font = BOLD
for j in range(3, 8): ws.cell(r, j, f'={L(j)}{R["in"]}+{L(j)}{R["out"]}').number_format = NUM; ws.cell(r, j).font = BOLD
ws.cell(r, 8, f'=H{R["appr"]}').number_format = NUM; ws.cell(r, 9, f'=I{R["pend"]}').number_format = NUM
ws.cell(r, 10, f'=SUM(C{r}:I{r})').number_format = NUM; ws.cell(r, 10).font = BOLD; r += 1
R['netp'] = r; ws.cell(r, 2, 'Net — procurement only (excluding PSR collections)')
for j in range(3, 8): ws.cell(r, j, f'={L(j)}{R["net"]}-{L(j)}{R["psr"]}').number_format = NUM
ws.cell(r, 8, f'=H{R["appr"]}').number_format = NUM; ws.cell(r, 9, f'=I{R["pend"]}').number_format = NUM; ws.cell(r, 10, f'=SUM(C{r}:I{r})').number_format = NUM; r += 1
for rr in range(7, r):
    for j in (R.get('in'), R.get('out')):
        pass
# totals column J for the by-month lines
for key in ('psr', 'proc', 'div', 'in', 'sup', 'lc', 'cus', 'out', 'oh'):
    rr = R[key]; ws.cell(rr, 10, f'=SUM(C{rr}:G{rr})').number_format = NUM
    if key in ('in', 'out'): ws.cell(rr, 10).font = BOLD
ws.cell(R['appr'], 10, f'=H{R["appr"]}').number_format = NUM; ws.cell(R['pend'], 10, f'=I{R["pend"]}').number_format = NUM
r += 1
ws.cell(r, 2, 'Reading: Oct-26 is closed (payments to 30/10, collections to 02/11). Nov-26 actuals fill as November payments and collections are pasted. Ongoing = open commitments carried into November; approved = in the Payment tracking log or on an approved list; the rest is pending approval. Overhead/corporate and PISTELLI/PELZ lines are kept in Cash OUT with Include? = No and shown as a memo.').font = Font(name=F, italic=True, size=9, color='595959')
ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10); ws.row_dimensions[r].height = 42
ws.column_dimensions['A'].width = 2; ws.column_dimensions['B'].width = 62
for c in 'CDEFGHIJ': ws.column_dimensions[c].width = 17
ws.freeze_panes = 'C7'

# reference tabs
for name, src_name, start in (('Tiering Master', 'Tiering Master', 4), ('OH Suppliers', 'OH Suppliers', 1), ('Log', 'Data Sources', 1)):
    ws = nb.create_sheet(name); src = wb[src_name]; first = True; n = 0
    for row in src.iter_rows(min_row=start, values_only=True):
        if not any(v not in (None, '') for v in row): continue
        ws.append([v for v in row]); n += 1
    style_header(ws, 1, ws.max_column)
    for c in range(1, ws.max_column + 1): ws.column_dimensions[L(c)].width = 22 if name != 'Log' else 40
    ws.freeze_panes = 'A2'
for ws in nb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F: c.font = Font(name=F, size=c.font.size if c.font and c.font.size else 10, bold=c.font.bold if c.font else False, italic=c.font.italic if c.font else False, color=c.font.color if c.font else None)
nb.calculation.fullCalcOnLoad = True; nb.save(OUT); print('written', OUT, round(os.path.getsize(OUT) / 1e6, 1), 'MB')
