"""Build the weekly 'Advances received from clients vs advances paid to suppliers' workbook.

Usage:
  python build_advances_report.py --ap <AP_Adv_payment.xlsx> --ar <AR_PRE.xlsx> --out <report.xlsx> \
      [--as-of "Week 3 Oct 2026"] [--ar-sheet 4421] [--ap-sheet "W1+W2 all items"] [--usd 9.0] [--xof 0.0165]

Sources
  AR side : GL 4421000000 'Advance received from Customer' (sheet '4421' of the AP file; Nature = Avance / Amrt d'avance)
  AP side : GL 3411000000 / 3411010000 advances to suppliers (sheet 'W1+W2 all items', header on row 11; Dr = paid, Cr = amortised)
  Check   : AR PRE register (sheet 'AR'); Status 1 = 'Advance' rows = unamortised balance (HT); TYPE/Invoice Type = advance invoices
All amounts written to the data sheets are raw values; every total in the report sheets is an Excel formula (SUMIFS).
"""
import argparse, re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo

W2_TO_LE = {'MA10': 'JESA', 'MA20': 'JETM', 'MA30': 'JEIN', 'MA90': 'JEST', 'SN10': 'JESE', 'US10': 'JEUS', 'CI10': 'JEWA'}
LE_ORDER = ['JESA', 'JETM', 'JEIN', 'JEST', 'JESE', 'JEUS', 'JEWA']
LE_CCY = {'JESA': 'MAD', 'JETM': 'MAD', 'JEIN': 'MAD', 'JEST': 'MAD', 'JESE': 'XOF', 'JEUS': 'USD', 'JEWA': 'XOF'}
# (GL BU code, report name, AR PRE 'Business Unit' label)
BU_TABLE = [
    ('FEC', 'Fertilizers & Chemicals', 'Fertilizers & Chemicals'),
    ('SP2M', 'SP2M Ad Hoc Program', 'SP2M Ad Hoc Program'),
    ('GMM', 'Global Mining', 'Global Mining'),
    ('ENE', 'Energy Operations (incl. PTX)', 'Energy Operations'),
    ('AMT', 'Asset Management (AM)', 'AM'),
    ('WEN', 'Water & Environment', 'Water & Environement'),
    ('BIN', 'Building & Infrastructure', 'Building & Infrastructure'),
    ('JET', 'JESA Technologies', 'JESA Technologies'),
    ('ADS', 'Advisory', 'Advisory'),
    ('M&I', 'M&I (legacy code)', ''),
    ('TRA', 'Transport (JEST)', ''),
    ('CORP', 'Corporate / Overhead (no BU)', ''),
    ('OTH', 'Other codes (SOC, JAC, JES)', ''),
    ('NA', '(Unassigned)', ''),
]
BU_CODES = [b[0] for b in BU_TABLE]
BU_NAME = {b[0]: b[1] for b in BU_TABLE}
BU_ALIAS = {'B&I': 'BIN', 'PTX': 'ENE', 'JESA': 'CORP', 'JETM': 'CORP', 'JEIN': 'CORP', 'JEUS': 'CORP', 'JEWA': 'CORP',
            'JEST': 'CORP', 'JESE': 'CORP', 'SOC': 'OTH', 'JAC': 'OTH', 'JES': 'OTH'}
PRE_BU_TO_CODE = {'fertilizers & chemicals': 'FEC', 'sp2m ad hoc program': 'SP2M', 'global mining': 'GMM',
                  'energy operations': 'ENE', 'am': 'AMT', 'water & environement': 'WEN', 'water & environment': 'WEN',
                  'building & infrastructure': 'BIN', 'jesa technologies': 'JET', 'advisory': 'ADS',
                  'building & capabilities': 'BIN'}

NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
FONT = 'Arial'
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=14, bold=True)
F_HDR = Font(name=FONT, size=10, bold=True, color='FFFFFF')
F_BLUE = Font(name=FONT, size=10, color='0000FF')
F_GREEN = Font(name=FONT, size=10, color='008000')
F_GREEN_B = Font(name=FONT, size=10, color='008000', bold=True)
F_BLACK_B = Font(name=FONT, size=10, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color='595959')
FILL_HDR = PatternFill('solid', start_color='1F3864')
FILL_SUB = PatternFill('solid', start_color='D9E1F2')
FILL_TOT = PatternFill('solid', start_color='BDD7EE')
FILL_YEL = PatternFill('solid', start_color='FFFF00')
FILL_GRP_AR = PatternFill('solid', start_color='2E75B6')
FILL_GRP_AP = PatternFill('solid', start_color='C55A11')
FILL_GRP_POS = PatternFill('solid', start_color='548235')
FILL_GRP_CHK = PatternFill('solid', start_color='7F6000')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
MAXR = 50000  # SUMIFS range depth on data sheets


def norm_bu(code):
    if code is None or (isinstance(code, float) and pd.isna(code)) or str(code).strip() == '':
        return 'NA'
    c = str(code).strip()
    c = BU_ALIAS.get(c, c)
    return c if c in BU_CODES else 'OTH'


NEW2OLD = {}  # new-format project code (QC.01458) -> legacy project root, learned from the AR PRE register


def root(x):
    """Project root: legacy 'Q36213P1.P71' -> 'Q36213P1'; 'Q82020P1-C1' -> 'Q82020P1'; new 'QC.01458.P53.S' -> 'QC.01458' (-> legacy root when unambiguous)."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return '(blank)'
    s = str(x).strip()
    if not s:
        return '(blank)'
    if s.upper().startswith('QC.'):
        key = '.'.join(s.split('.')[:2])
        return NEW2OLD.get(key, key)
    return re.split(r'[.\-]', s)[0]


def load_ar_gl(path, sheet):
    d = pd.read_excel(path, sheet_name=sheet)
    d.columns = [str(c).strip() for c in d.columns]
    d = d[d['G/L Account'].astype(str).str.startswith('4421')].copy()
    cc = d['Company Code'].astype(str).str.strip()
    d['LE'] = cc.map(lambda c: W2_TO_LE.get(c, c))
    d['BUc'] = d['BU'].map(norm_bu)
    src = d['WBS'] if 'WBS' in d else d['Assignment']
    d['Root'] = src.map(root)
    d['Nature'] = d['Nature'].astype(str).str.strip()
    d['Amount'] = pd.to_numeric(d['Company Code Currency Value'], errors='coerce').fillna(0.0)
    return d


def load_ap_gl(path, sheet):
    d = pd.read_excel(path, sheet_name=sheet, header=10)
    d.columns = [str(c).strip() for c in d.columns]
    d = d[d['G/L Account'].astype(str).str.startswith('3411')].copy()
    d['LE'] = d['Entité (W1)'].astype(str).str.strip()
    d['BUc'] = d['BU'].map(norm_bu)
    d['Root'] = d['WBS'].map(root)
    d['Dr'] = pd.to_numeric(d['Dr'], errors='coerce').fillna(0.0)
    d['Cr'] = pd.to_numeric(d['Cr'], errors='coerce').fillna(0.0)
    d['Amount'] = pd.to_numeric(d['Company Code Currency Value'], errors='coerce').fillna(0.0)
    d['Category'] = d['Category'].fillna('Project').astype(str).str.strip()
    d['Ccy'] = d['Company Code Currency Key'].astype(str).str.strip()
    return d


def load_ar_pre(path, sheet):
    d = pd.read_excel(path, sheet_name=sheet)
    d.columns = [str(c).strip() for c in d.columns]
    s1 = d['Status 1'].astype(str).str.strip().str.lower()
    typ = d['TYPE'].astype(str).str.strip().str.lower()
    it = d['Invoice Type'].astype(str).str.strip().str.lower()
    is_bal = s1.eq('advance')
    is_pend = s1.eq('advance payment')
    is_inv = ((typ.eq('advance')) | it.str.contains('advance|avance|acompte', regex=True)) & ~is_bal & ~is_pend
    # learn new-code -> legacy-root mapping (only where unambiguous)
    mp = d[['Project', 'New Project']].dropna().copy()
    mp['old'] = mp['Project'].astype(str).str.strip().map(lambda v: re.split(r'[.\-]', v)[0])
    mp['new'] = mp['New Project'].astype(str).str.strip().map(lambda v: '.'.join(v.split('.')[:2]))
    mp['bal'] = is_bal.loc[mp.index].astype(int)
    # a new code that spans several legacy roots is mapped to the legacy root carrying the advance balance, else the most frequent one
    cnt = mp.groupby(['new', 'old']).agg(bal=('bal', 'sum'), n=('old', 'size')).reset_index().sort_values(['new', 'bal', 'n'], ascending=[True, False, False])
    best = cnt.drop_duplicates('new').set_index('new')['old']
    NEW2OLD.update({k: v for k, v in best.items() if k.upper().startswith('QC.')})
    d['RowType'] = ''
    d.loc[is_inv, 'RowType'] = 'Advance invoice'
    d.loc[is_pend, 'RowType'] = 'Advance pending PO (not invoiced)'
    d.loc[is_bal, 'RowType'] = 'Unamortised balance'
    d = d[d['RowType'] != ''].copy()
    le = d['LE'].fillna('').astype(str).str.strip().str.upper().replace({'NAN': '', 'NONE': ''})
    d['LEn'] = le.map(lambda c: W2_TO_LE.get(c, c if c else ''))
    d['Root'] = d['Project'].map(root)
    # blank LE: take the LE used by other lines of the same project root
    known = d[d['LEn'] != ''].groupby('Root')['LEn'].agg(lambda s: s.value_counts().index[0])
    d['LEn'] = [v if v else known.get(rt, '(blank)') for v, rt in zip(d['LEn'], d['Root'])]
    d['BUc'] = d['Business Unit'].astype(str).str.strip().str.lower().map(lambda s: PRE_BU_TO_CODE.get(s, 'NA'))
    d['HT'] = pd.to_numeric(d['Total Amount HT'], errors='coerce').fillna(0.0)
    d['TTC'] = pd.to_numeric(d['Total Amount TTC'], errors='coerce').fillna(0.0)
    return d


def style_header(ws, row, ncols, fill=FILL_HDR):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HDR; cell.fill = fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def put(ws, r, c, v, font=F_BASE, fmt=None, fill=None, align=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font
    if fmt: cell.number_format = fmt
    if fill: cell.fill = fill
    if align: cell.alignment = align
    cell.border = BORDER
    return cell


def write_df(ws, df, cols, start_row=1, fmts=None):
    """cols: list of (header, series_or_key). Writes header at start_row, data below. Returns last row."""
    fmts = fmts or {}
    for j, (h, _) in enumerate(cols, 1):
        ws.cell(row=start_row, column=j, value=h)
    style_header(ws, start_row, len(cols))
    data = [df[k].tolist() if isinstance(k, str) else k for _, k in cols]
    n = len(df)
    for i in range(n):
        r = start_row + 1 + i
        for j in range(len(cols)):
            v = data[j][i]
            if isinstance(v, float) and pd.isna(v): v = None
            if isinstance(v, pd.Timestamp): v = v.to_pydatetime()
            cell = ws.cell(row=r, column=j + 1, value=v)
            cell.font = F_BASE
            f = fmts.get(cols[j][0])
            if f: cell.number_format = f
    return start_row + n


def build(ap_path, ar_path, out, as_of, ar_sheet, ap_sheet, pre_sheet, usd, xof):
    pre = load_ar_pre(ar_path, pre_sheet)   # first: learns the new-code -> legacy project root map
    argl = load_ar_gl(ap_path, ar_sheet)
    apgl = load_ap_gl(ap_path, ap_sheet)

    wb = Workbook()
    # ------------------------------------------------------------------ Inputs
    wi = wb.active; wi.title = 'Inputs'
    put(wi, 1, 1, 'Inputs & mappings', F_TITLE)
    put(wi, 3, 1, 'Report as of', F_BOLD); put(wi, 3, 2, as_of, F_BLUE, fill=FILL_YEL)
    put(wi, 4, 1, 'AP category shown on Summary / By BU', F_BOLD); put(wi, 4, 2, 'All', F_BLUE, fill=FILL_YEL)
    put(wi, 4, 3, 'Type: All / Project / Overhead/Annulation', F_NOTE)
    put(wi, 6, 1, 'FX rates to MAD', F_BOLD)
    for j, h in enumerate(['Currency', 'Rate to MAD', 'Source / note'], 1): put(wi, 7, j, h)
    style_header(wi, 7, 3)
    fx = [('MAD', 1.0, 'Reporting currency'), ('USD', usd, 'Source: ASSUMPTION - update to Bank Al-Maghrib closing rate at report date'),
          ('XOF', xof, 'Source: ASSUMPTION (EUR peg 655.957 XOF) - update to closing rate at report date')]
    for i, (c, r, n) in enumerate(fx):
        put(wi, 8 + i, 1, c); put(wi, 8 + i, 2, r, F_BLUE, '0.0000', FILL_YEL if c != 'MAD' else None); put(wi, 8 + i, 3, n, F_NOTE)
    FX_RANGE = "Inputs!$A$8:$B$10"
    put(wi, 13, 1, 'Legal entities', F_BOLD)
    for j, h in enumerate(['W2 code', 'LE (W1)', 'Local currency'], 1): put(wi, 14, j, h)
    style_header(wi, 14, 3)
    for i, (w2, le) in enumerate(W2_TO_LE.items()):
        put(wi, 15 + i, 1, w2); put(wi, 15 + i, 2, le); put(wi, 15 + i, 3, LE_CCY[le])
    r0 = 15 + len(W2_TO_LE) + 2
    put(wi, r0, 1, 'Business units (GL code -> report name -> AR PRE label)', F_BOLD)
    for j, h in enumerate(['BU code (GL)', 'Report name', 'AR PRE "Business Unit" label', 'Also mapped from'], 1): put(wi, r0 + 1, j, h)
    style_header(wi, r0 + 1, 4)
    also = {}
    for k, v in BU_ALIAS.items(): also.setdefault(v, []).append(k)
    for i, (c, n, p) in enumerate(BU_TABLE):
        rr = r0 + 2 + i
        put(wi, rr, 1, c); put(wi, rr, 2, n); put(wi, rr, 3, p); put(wi, rr, 4, ', '.join(also.get(c, [])))
    BU_RANGE = f"Inputs!$A${r0 + 2}:$B${r0 + 1 + len(BU_TABLE)}"
    widths(wi, [38, 30, 70, 30])

    # ------------------------------------------------------------------ AR_GL data
    wa = wb.create_sheet('AR_GL')
    argl = argl.sort_values(['LE', 'BUc', 'Root', 'Posting Date'])
    cols = [('Company Code', 'Company Code'), ('LE', 'LE'), ('Document Number', 'Document Number'), ('Doc type', 'Document type'),
            ('Posting Date', 'Posting Date'), ('Nature', 'Nature'), ('Status', 'Status'), ('Customer', 'Customer Account: Name 1'),
            ('OCP/NON OCP', 'OCP/NON OCP'), ('Project root', 'Root'), ('WBS', 'Assignment'), ('BU (raw)', 'BU'), ('BU', 'BUc'),
            ('Text', 'Text'), ('Amount GL (MAD)', 'Amount')]
    last = write_df(wa, argl, cols, fmts={'Amount GL (MAD)': NUM, 'Posting Date': 'dd-mmm-yyyy'})
    for j, h in enumerate(['Advance received (MAD)', 'Amortisation (MAD)', 'Net outstanding (MAD)'], 16):
        wa.cell(row=1, column=j, value=h)
    style_header(wa, 1, 18)
    for r in range(2, last + 1):
        put(wa, r, 16, f'=IF($F{r}="Avance",-$O{r},0)', fmt=NUM)
        put(wa, r, 17, f'=IF($F{r}="Amrt d\'avance",$O{r},0)', fmt=NUM)
        put(wa, r, 18, f'=P{r}-Q{r}', fmt=NUM)
    wa.freeze_panes = 'A2'; wa.auto_filter.ref = f'A1:R{last}'
    widths(wa, [10, 7, 13, 7, 12, 13, 7, 28, 10, 12, 14, 8, 6, 40, 16, 16, 16, 16])
    wa.cell(row=1, column=15).comment = Comment("Source: SAP GL 4421000000 'Advance received from Customer', sheet '4421' of AP_Adv_payment file. Credit (negative) = advance received; debit (positive) = amortisation.", 'Claude')

    # ------------------------------------------------------------------ AP_GL data
    wp = wb.create_sheet('AP_GL')
    apgl = apgl.sort_values(['LE', 'Category', 'BUc', 'Root', 'Posting Date'])
    cols = [('LE', 'LE'), ('W2 code', 'Entité (W2)'), ('G/L Account', 'G/L Account'), ('Doc type', 'Document type'),
            ('Document Number', 'Document Number'), ('Posting Date', 'Posting Date'), ('Supplier', 'Vendor Account: Name 1'),
            ('Text', 'Text'), ('Category', 'Category'), ('Profit center', 'BU/Pofit center ME2L'), ('BU (raw)', 'BU'), ('BU', 'BUc'),
            ('Client', 'Client'), ('Project root', 'Root'), ('Ccy', 'Ccy'), ('Dr (local)', 'Dr'), ('Cr (local)', 'Cr'), ('Net (local)', 'Amount')]
    last_p = write_df(wp, apgl, cols, fmts={'Dr (local)': NUM, 'Cr (local)': NUM, 'Net (local)': NUM, 'Posting Date': 'dd-mmm-yyyy'})
    for j, h in enumerate(['FX to MAD', 'Advance paid (MAD)', 'Amortisation (MAD)', 'Net outstanding (MAD)'], 19):
        wp.cell(row=1, column=j, value=h)
    style_header(wp, 1, 22)
    for r in range(2, last_p + 1):
        put(wp, r, 19, f'=IFERROR(VLOOKUP($O{r},{FX_RANGE},2,FALSE),1)', F_GREEN, '0.0000')
        put(wp, r, 20, f'=$P{r}*$S{r}', fmt=NUM)
        put(wp, r, 21, f'=-$Q{r}*$S{r}', fmt=NUM)
        put(wp, r, 22, f'=T{r}-U{r}', fmt=NUM)
    wp.freeze_panes = 'A2'; wp.auto_filter.ref = f'A1:V{last_p}'
    widths(wp, [7, 8, 12, 7, 13, 12, 30, 40, 18, 12, 8, 6, 24, 12, 5, 15, 15, 15, 8, 16, 16, 16])
    wp.cell(row=1, column=16).comment = Comment("Source: SAP GL 3411000000 / 3411010000 advances to suppliers, sheet 'W1+W2 all items' of AP_Adv_payment file. Dr = advance paid, Cr = amortisation / recovery on supplier invoices.", 'Claude')

    # ------------------------------------------------------------------ AR_PRE data
    wr = wb.create_sheet('AR_PRE')
    pre = pre.sort_values(['LEn', 'BUc', 'Root', 'RowType'])
    cols = [('LE', 'LEn'), ('LE (raw)', 'LE'), ('Business Unit (PRE)', 'Business Unit'), ('BU', 'BUc'), ('Project root', 'Root'),
            ('Project', 'Project'), ('Project Name', 'Project Name'), ('Client', 'Client 1'), ('Invoice Number', 'Invoice Number'),
            ('Invoice Date', 'Invoice Date'), ('Status 1', 'Status 1'), ('Status 2', 'Status 2'), ('TYPE', 'TYPE'),
            ('Invoice Type', 'Invoice Type'), ('Row type', 'RowType'), ('Amount HT', 'HT'), ('Amount TTC', 'TTC'), ('Comments PA', 'Comments PA')]
    last_r = write_df(wr, pre, cols, fmts={'Amount HT': NUM, 'Amount TTC': NUM, 'Invoice Date': 'dd-mmm-yyyy'})
    for j, h in enumerate(['Unamortised balance HT (MAD)', 'Advance invoiced HT (MAD)'], 19):
        wr.cell(row=1, column=j, value=h)
    style_header(wr, 1, 20)
    for r in range(2, last_r + 1):
        put(wr, r, 19, f'=IF($O{r}="Unamortised balance",-$P{r},0)', fmt=NUM)
        put(wr, r, 20, f'=IF($O{r}="Advance invoice",$P{r},0)', fmt=NUM)
    wr.freeze_panes = 'A2'; wr.auto_filter.ref = f'A1:T{last_r}'
    widths(wr, [7, 7, 22, 6, 11, 11, 34, 20, 12, 12, 14, 30, 22, 16, 22, 15, 15, 30, 18, 18])
    wr.cell(row=1, column=15).comment = Comment("Row type: 'Unamortised balance' = AR PRE lines with Status 1 = Advance (negative HT = advance still to be amortised against billing); 'Advance invoice' = advance invoices issued (TYPE / Invoice Type = Advance); 'Advance pending PO' = advance not yet invoiced (excluded from checks).", 'Claude')

    # ------------------------------------------------------------------ helper strings
    AR = lambda col: f"AR_GL!${col}$2:${col}${MAXR}"
    AP = lambda col: f"AP_GL!${col}$2:${col}${MAXR}"
    PR = lambda col: f"AR_PRE!${col}$2:${col}${MAXR}"
    CAT = 'IF(Inputs!$B$4="All","*",Inputs!$B$4)'

    # ------------------------------------------------------------------ Summary (LE x BU)
    ws = wb.create_sheet('Summary', 0)
    put(ws, 1, 1, 'Advances received from clients vs advances paid to suppliers - by Legal Entity and Business Unit', F_TITLE)
    put(ws, 2, 1, '="As of: "&Inputs!$B$3&"   |   Amounts in MAD   |   AP category: "&Inputs!$B$4', F_GREEN)
    put(ws, 3, 1, 'Gross = advances before amortisation ("excluding amortisation"). Net = gross less amortisation = outstanding balance ("including amortisation").', F_NOTE)
    hdr_groups = [(4, 6, 'AR - Advances received from clients (GL 4421)', FILL_GRP_AR), (7, 9, 'AP - Advances paid to suppliers (GL 3411)', FILL_GRP_AP),
                  (10, 12, 'Position AR vs AP', FILL_GRP_POS)]
    for c1, c2, t, f in hdr_groups:
        ws.merge_cells(start_row=5, start_column=c1, end_row=5, end_column=c2)
        cell = ws.cell(row=5, column=c1, value=t); cell.font = F_HDR; cell.fill = f; cell.alignment = Alignment(horizontal='center')
    heads = ['LE', 'BU', 'BU name', 'Received (gross)', 'Amortised', 'Net outstanding', 'Paid (gross)', 'Amortised', 'Net outstanding',
             'Net: AR - AP', 'Gross: AR - AP', 'AP net / AR net']
    for j, h in enumerate(heads, 1): ws.cell(row=6, column=j, value=h)
    style_header(ws, 6, len(heads))
    pairs = set(zip(argl['LE'], argl['BUc'])) | set(zip(apgl['LE'], apgl['BUc']))
    r = 7
    sub_rows = []
    for le in LE_ORDER + sorted({p[0] for p in pairs} - set(LE_ORDER)):
        bus = [b for b in BU_CODES if (le, b) in pairs]
        if not bus: continue
        first = r
        for b in bus:
            put(ws, r, 1, le); put(ws, r, 2, b); put(ws, r, 3, f'=IFERROR(VLOOKUP($B{r},{BU_RANGE},2,FALSE),"")', F_GREEN)
            put(ws, r, 4, f'=SUMIFS({AR("P")},{AR("B")},$A{r},{AR("M")},$B{r})', F_GREEN, NUM)
            put(ws, r, 5, f'=SUMIFS({AR("Q")},{AR("B")},$A{r},{AR("M")},$B{r})', F_GREEN, NUM)
            put(ws, r, 6, f'=D{r}-E{r}', fmt=NUM)
            put(ws, r, 7, f'=SUMIFS({AP("T")},{AP("A")},$A{r},{AP("L")},$B{r},{AP("I")},{CAT})', F_GREEN, NUM)
            put(ws, r, 8, f'=SUMIFS({AP("U")},{AP("A")},$A{r},{AP("L")},$B{r},{AP("I")},{CAT})', F_GREEN, NUM)
            put(ws, r, 9, f'=G{r}-H{r}', fmt=NUM)
            put(ws, r, 10, f'=F{r}-I{r}', fmt=NUM); put(ws, r, 11, f'=D{r}-G{r}', fmt=NUM)
            put(ws, r, 12, f'=IF(F{r}=0,"",I{r}/F{r})', fmt=PCT)
            r += 1
        put(ws, r, 1, le, F_BOLD, fill=FILL_SUB); put(ws, r, 2, 'Total', F_BOLD, fill=FILL_SUB); put(ws, r, 3, f'="Total "&A{r}', F_BOLD, fill=FILL_SUB)
        for c in range(4, 12):
            put(ws, r, c, f'=SUBTOTAL(9,{get_column_letter(c)}{first}:{get_column_letter(c)}{r - 1})', F_BOLD, NUM, FILL_SUB)
        put(ws, r, 12, f'=IF(F{r}=0,"",I{r}/F{r})', F_BOLD, PCT, FILL_SUB)
        sub_rows.append(r); r += 1
    put(ws, r, 1, 'Grand total', F_BOLD, fill=FILL_TOT); put(ws, r, 2, '', fill=FILL_TOT); put(ws, r, 3, 'All LEs / all BUs', F_BOLD, fill=FILL_TOT)
    for c in range(4, 12):
        L = get_column_letter(c)
        put(ws, r, c, '=' + '+'.join(f'{L}{s}' for s in sub_rows), F_BOLD, NUM, FILL_TOT)
    put(ws, r, 12, f'=IF(F{r}=0,"",I{r}/F{r})', F_BOLD, PCT, FILL_TOT)
    gt = r
    r += 2
    put(ws, r, 1, 'Control totals (must be zero)', F_BOLD); r += 1
    put(ws, r, 1, 'AR: grand total net vs sum of AR_GL sheet'); put(ws, r, 4, f'=F{gt}-SUM({AR("R")})', F_GREEN, NUM); r += 1
    put(ws, r, 1, 'AP: grand total net vs sum of AP_GL sheet (only when category = All)')
    put(ws, r, 4, f'=IF(Inputs!$B$4="All",I{gt}-SUM({AP("V")}),"n/a")', F_GREEN, NUM); r += 1
    put(ws, r, 1, 'Reading: a positive "Net: AR - AP" means the LE/BU holds more client advance (unamortised) than it has advanced to its suppliers.', F_NOTE)
    r += 1
    put(ws, r, 1, 'AR GL extract currently covers JESA (MA10) only. JETM / JEUS / JEIN client advances appear in the AR PRE check sheet but not here - request GL 4421 for MA20/US10/MA30.', F_NOTE)
    ws.freeze_panes = 'D7'
    widths(ws, [10, 7, 30, 17, 15, 17, 17, 15, 17, 17, 17, 13])

    # ------------------------------------------------------------------ By LE
    wl = wb.create_sheet('By LE', 1)
    put(wl, 1, 1, 'Summary by Legal Entity', F_TITLE)
    put(wl, 2, 1, '="As of: "&Inputs!$B$3&"   |   Amounts in MAD unless stated"', F_GREEN)
    grp = [(3, 5, 'AR - received from clients (GL 4421)', FILL_GRP_AR), (6, 8, 'AP - paid to suppliers: Project', FILL_GRP_AP),
           (9, 11, 'AP - paid to suppliers: Overhead/Annulation', FILL_GRP_AP), (12, 14, 'AP - total', FILL_GRP_AP),
           (15, 16, 'Position', FILL_GRP_POS), (17, 20, 'AP tie-out to TB (local currency)', FILL_GRP_CHK)]
    for c1, c2, t, f in grp:
        wl.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
        cell = wl.cell(row=4, column=c1, value=t); cell.font = F_HDR; cell.fill = f; cell.alignment = Alignment(horizontal='center', wrap_text=True)
    heads = ['LE', 'Ccy', 'Received (gross)', 'Amortised', 'Net', 'Paid (gross)', 'Amortised', 'Net', 'Paid (gross)', 'Amortised', 'Net',
             'Paid (gross)', 'Amortised', 'Net', 'Net: AR - AP', 'Gross: AR - AP', 'AP net (local)', 'TB W2 balance (local)', 'Difference', 'TB source']
    for j, h in enumerate(heads, 1): wl.cell(row=5, column=j, value=h)
    style_header(wl, 5, len(heads))
    tb = {'JEWA': 45900593.75, 'JESA': 2895388325.54, 'JETM': 361720697.12, 'JEIN': 75605.45, 'JEST': 1104119.88, 'JESE': 3100000, 'JEUS': 8533694.24}
    r = 6
    for le in LE_ORDER:
        put(wl, r, 1, le); put(wl, r, 2, LE_CCY.get(le, ''))
        put(wl, r, 3, f'=SUMIFS({AR("P")},{AR("B")},$A{r})', F_GREEN, NUM); put(wl, r, 4, f'=SUMIFS({AR("Q")},{AR("B")},$A{r})', F_GREEN, NUM)
        put(wl, r, 5, f'=C{r}-D{r}', fmt=NUM)
        for base, cat in ((6, '"Project"'), (9, '"Overhead/Annulation"')):
            put(wl, r, base, f'=SUMIFS({AP("T")},{AP("A")},$A{r},{AP("I")},{cat})', F_GREEN, NUM)
            put(wl, r, base + 1, f'=SUMIFS({AP("U")},{AP("A")},$A{r},{AP("I")},{cat})', F_GREEN, NUM)
            put(wl, r, base + 2, f'={get_column_letter(base)}{r}-{get_column_letter(base + 1)}{r}', fmt=NUM)
        put(wl, r, 12, f'=F{r}+I{r}', fmt=NUM); put(wl, r, 13, f'=G{r}+J{r}', fmt=NUM); put(wl, r, 14, f'=L{r}-M{r}', fmt=NUM)
        put(wl, r, 15, f'=E{r}-N{r}', fmt=NUM); put(wl, r, 16, f'=C{r}-L{r}', fmt=NUM)
        put(wl, r, 17, f'=SUMIFS({AP("R")},{AP("A")},$A{r})', F_GREEN, NUM)
        put(wl, r, 18, tb.get(le), F_BLUE, NUM); put(wl, r, 19, f'=Q{r}-R{r}', fmt=NUM)
        put(wl, r, 20, "Source: AP_Adv_payment_Oct_5th.xlsx, sheet 'W1+W2 all items', col D 'TB W2' (rows 2-8)", F_NOTE)
        r += 1
    put(wl, r, 1, 'Total (MAD)', F_BOLD, fill=FILL_TOT); put(wl, r, 2, '', fill=FILL_TOT)
    for c in range(3, 17):
        L = get_column_letter(c); put(wl, r, c, f'=SUM({L}6:{L}{r - 1})', F_BOLD, NUM, FILL_TOT)
    for c in range(17, 21): put(wl, r, c, '', fill=FILL_TOT)
    r += 2
    put(wl, r, 1, 'Notes', F_BOLD); r += 1
    for n in ['AR GL 4421 extract covers JESA only (company codes JESA/MA10). Other LEs show zero on the AR side until their GL 4421 extract is added.',
              'JEUS books in USD, JESE and JEWA in XOF: MAD columns use the FX rates on the Inputs sheet (yellow cells - update weekly).',
              'TB W2 balance is the trial-balance figure typed in the source file; difference should be zero when the GL line items are complete.',
              '"Overhead/Annulation" = non-project supplier advances (deposits, overhead POs) and cancellations, as categorised in the source file.']:
        put(wl, r, 1, n, F_NOTE); r += 1
    wl.freeze_panes = 'C6'
    widths(wl, [12, 5, 16, 14, 16, 16, 14, 16, 14, 13, 14, 16, 14, 16, 16, 16, 16, 18, 14, 60])

    # ------------------------------------------------------------------ By BU
    wbu = wb.create_sheet('By BU', 2)
    put(wbu, 1, 1, 'Summary by Business Unit (all legal entities)', F_TITLE)
    put(wbu, 2, 1, '="As of: "&Inputs!$B$3&"   |   Amounts in MAD   |   AP category: "&Inputs!$B$4', F_GREEN)
    for c1, c2, t, f in [(3, 5, 'AR - received from clients (GL 4421)', FILL_GRP_AR), (6, 8, 'AP - paid to suppliers (GL 3411)', FILL_GRP_AP), (9, 11, 'Position', FILL_GRP_POS)]:
        wbu.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
        cell = wbu.cell(row=4, column=c1, value=t); cell.font = F_HDR; cell.fill = f; cell.alignment = Alignment(horizontal='center')
    heads = ['BU', 'BU name', 'Received (gross)', 'Amortised', 'Net', 'Paid (gross)', 'Amortised', 'Net', 'Net: AR - AP', 'Gross: AR - AP', 'AP net / AR net']
    for j, h in enumerate(heads, 1): wbu.cell(row=5, column=j, value=h)
    style_header(wbu, 5, len(heads))
    r = 6
    used_bus = [b for b in BU_CODES if b in set(argl['BUc']) | set(apgl['BUc'])]
    for b in used_bus:
        put(wbu, r, 1, b); put(wbu, r, 2, f'=IFERROR(VLOOKUP($A{r},{BU_RANGE},2,FALSE),"")', F_GREEN)
        put(wbu, r, 3, f'=SUMIFS({AR("P")},{AR("M")},$A{r})', F_GREEN, NUM); put(wbu, r, 4, f'=SUMIFS({AR("Q")},{AR("M")},$A{r})', F_GREEN, NUM)
        put(wbu, r, 5, f'=C{r}-D{r}', fmt=NUM)
        put(wbu, r, 6, f'=SUMIFS({AP("T")},{AP("L")},$A{r},{AP("I")},{CAT})', F_GREEN, NUM)
        put(wbu, r, 7, f'=SUMIFS({AP("U")},{AP("L")},$A{r},{AP("I")},{CAT})', F_GREEN, NUM)
        put(wbu, r, 8, f'=F{r}-G{r}', fmt=NUM); put(wbu, r, 9, f'=E{r}-H{r}', fmt=NUM); put(wbu, r, 10, f'=C{r}-F{r}', fmt=NUM)
        put(wbu, r, 11, f'=IF(E{r}=0,"",H{r}/E{r})', fmt=PCT); r += 1
    put(wbu, r, 1, 'Total', F_BOLD, fill=FILL_TOT); put(wbu, r, 2, '', fill=FILL_TOT)
    for c in range(3, 11):
        L = get_column_letter(c); put(wbu, r, c, f'=SUM({L}6:{L}{r - 1})', F_BOLD, NUM, FILL_TOT)
    put(wbu, r, 11, f'=IF(E{r}=0,"",H{r}/E{r})', F_BOLD, PCT, FILL_TOT)
    wbu.freeze_panes = 'C6'
    widths(wbu, [8, 30, 17, 15, 17, 17, 15, 17, 17, 17, 13])

    # ------------------------------------------------------------------ AR PRE check (LE x BU)
    wc = wb.create_sheet('AR PRE check', 3)
    put(wc, 1, 1, 'AR check: GL 4421 (accounting) vs AR PRE register (project accountants) - by LE and BU, amounts HT in MAD', F_TITLE)
    put(wc, 2, 1, '="As of: "&Inputs!$B$3', F_GREEN)
    put(wc, 3, 1, 'Unamortised balance = client advance received and not yet deducted from billing. GL: net of GL 4421 (JESA only). PRE: lines with Status 1 = "Advance" (sign flipped).', F_NOTE)
    for c1, c2, t, f in [(3, 5, 'Unamortised balance (net)', FILL_GRP_CHK), (6, 8, 'Gross advances invoiced / received', FILL_GRP_CHK), (9, 9, 'PRE memo', FILL_GRP_CHK)]:
        wc.merge_cells(start_row=4, start_column=c1, end_row=4, end_column=c2)
        cell = wc.cell(row=4, column=c1, value=t); cell.font = F_HDR; cell.fill = f; cell.alignment = Alignment(horizontal='center')
    heads = ['LE', 'BU', 'GL 4421 net', 'AR PRE balance', 'Variance GL - PRE', 'GL received (gross)', 'PRE advance invoices', 'Variance GL - PRE', 'PRE advances pending PO']
    for j, h in enumerate(heads, 1): wc.cell(row=5, column=j, value=h)
    style_header(wc, 5, len(heads))
    pairs_c = set(zip(argl['LE'], argl['BUc'])) | set(zip(pre['LEn'], pre['BUc']))
    r = 6; subs = []
    for le in LE_ORDER + sorted({p[0] for p in pairs_c} - set(LE_ORDER)):
        bus = [b for b in BU_CODES if (le, b) in pairs_c]
        if not bus: continue
        first = r
        for b in bus:
            put(wc, r, 1, le); put(wc, r, 2, b)
            put(wc, r, 3, f'=SUMIFS({AR("R")},{AR("B")},$A{r},{AR("M")},$B{r})', F_GREEN, NUM)
            put(wc, r, 4, f'=SUMIFS({PR("S")},{PR("A")},$A{r},{PR("D")},$B{r})', F_GREEN, NUM)
            put(wc, r, 5, f'=C{r}-D{r}', fmt=NUM)
            put(wc, r, 6, f'=SUMIFS({AR("P")},{AR("B")},$A{r},{AR("M")},$B{r})', F_GREEN, NUM)
            put(wc, r, 7, f'=SUMIFS({PR("T")},{PR("A")},$A{r},{PR("D")},$B{r})', F_GREEN, NUM)
            put(wc, r, 8, f'=F{r}-G{r}', fmt=NUM)
            put(wc, r, 9, f'=SUMIFS({PR("P")},{PR("A")},$A{r},{PR("D")},$B{r},{PR("O")},"Advance pending PO*")', F_GREEN, NUM)
            r += 1
        put(wc, r, 1, le, F_BOLD, fill=FILL_SUB); put(wc, r, 2, 'Total', F_BOLD, fill=FILL_SUB)
        for c in range(3, 10):
            L = get_column_letter(c); put(wc, r, c, f'=SUBTOTAL(9,{L}{first}:{L}{r - 1})', F_BOLD, NUM, FILL_SUB)
        subs.append(r); r += 1
    put(wc, r, 1, 'Grand total', F_BOLD, fill=FILL_TOT); put(wc, r, 2, '', fill=FILL_TOT)
    for c in range(3, 10):
        L = get_column_letter(c); put(wc, r, c, '=' + '+'.join(f'{L}{s}' for s in subs), F_BOLD, NUM, FILL_TOT)
    r += 2
    for n in ['How to read: a zero variance on the unamortised balance confirms the GL AR figure. Variances are timing (amortisation booked in one source only), VAT/HT differences, or project coding differences - see sheet "PRE vs GL by project".',
              'Gross comparison is indicative only: AR PRE tags advance invoices by invoice type, which is less complete than the GL "Avance" nature (e.g. SP2M advances are tagged as Procurement/Subcontract in PRE).',
              'LEs other than JESA have no GL 4421 lines in the current extract, so their GL columns are zero by construction, not because the balance is nil.']:
        put(wc, r, 1, n, F_NOTE); r += 1
    wc.freeze_panes = 'C6'
    widths(wc, [10, 7, 17, 17, 17, 18, 18, 17, 20])

    # ------------------------------------------------------------------ PRE vs GL by project
    wj = wb.create_sheet('PRE vs GL by project', 4)
    put(wj, 1, 1, 'Unamortised client advance by project root: GL 4421 vs AR PRE (HT, MAD) - sorted by absolute variance', F_TITLE)
    put(wj, 2, 1, 'Project roots present in either source. Filter on column F to isolate projects to investigate.', F_NOTE)
    heads = ['Project root', 'LE (GL)', 'BU (GL)', 'GL 4421 net', 'AR PRE balance', 'Variance GL - PRE', 'Abs variance', 'Project name (PRE)', 'LE (PRE)', 'BU (PRE)']
    for j, h in enumerate(heads, 1): wj.cell(row=4, column=j, value=h)
    style_header(wj, 4, len(heads))
    g_net = argl.groupby('Root')['Amount'].sum().mul(-1)
    p_bal = pre[pre['RowType'] == 'Unamortised balance'].groupby('Root')['HT'].sum().mul(-1)
    roots = sorted(set(g_net.index) | set(p_bal.index), key=lambda x: -abs(g_net.get(x, 0) - p_bal.get(x, 0)))
    gl_le = argl.groupby('Root')['LE'].agg(lambda s: s.value_counts().index[0]); gl_bu = argl.groupby('Root')['BUc'].agg(lambda s: s.value_counts().index[0])
    p_name = pre.groupby('Root')['Project Name'].agg(lambda s: s.dropna().astype(str).iloc[0] if s.dropna().size else '')
    p_le = pre.groupby('Root')['LEn'].agg(lambda s: s.value_counts().index[0]); p_bu = pre.groupby('Root')['BUc'].agg(lambda s: s.value_counts().index[0])
    r = 5
    for x in roots:
        put(wj, r, 1, x); put(wj, r, 2, gl_le.get(x, '')); put(wj, r, 3, gl_bu.get(x, '')); put(wj, r, 9, p_le.get(x, '')); put(wj, r, 10, p_bu.get(x, ''))
        put(wj, r, 4, f'=SUMIFS({AR("R")},{AR("J")},$A{r})', F_GREEN, NUM)
        put(wj, r, 5, f'=SUMIFS({PR("S")},{PR("E")},$A{r})', F_GREEN, NUM)
        put(wj, r, 6, f'=D{r}-E{r}', fmt=NUM); put(wj, r, 7, f'=ABS(F{r})', fmt=NUM); put(wj, r, 8, p_name.get(x, ''))
        r += 1
    put(wj, r, 1, 'Total', F_BOLD, fill=FILL_TOT)
    for c in range(4, 8):
        L = get_column_letter(c); put(wj, r, c, f'=SUM({L}5:{L}{r - 1})', F_BOLD, NUM, FILL_TOT)
    wj.freeze_panes = 'B5'; wj.auto_filter.ref = f'A4:J{r - 1}'
    widths(wj, [14, 8, 8, 17, 17, 17, 15, 45, 8, 8])

    # ------------------------------------------------------------------ Version check + README
    wv = wb.create_sheet('Version check', 5)
    put(wv, 1, 1, 'Which GL 4421 version is used, and why', F_TITLE)
    rows = [
        ('Item', "Sheet 'Data W1' (older version)", "Sheet '4421' (version used)", "Pivot 'TCD (2)' (cached)"),
        ('Line count', 4913, 4272, ''),
        ('Latest entry date in extract', '27-Apr-2026', '29-Sep-2026', ''),
        ('Latest posting date', '24-Apr-2026', '25-Sep-2026', ''),
        ('Company codes', 'JESA', 'JESA + MA10 (same LE, W2 code)', ''),
        ('Net balance GL 4421 (MAD)', -7638485184.75, -7254266963.03, -7254172762.94),
        ('Enriched with Nature / Status / BU / project root', 'No', 'Yes', ''),
        ('Lines posted May-Sep 2026 (only in 4421)', '', 298882480.49, ''),
        ('Lines only in Data W1 (cleared pairs, AB reversals, duplicates)', -86092407.90, '', ''),
        ('Same documents, different amounts', -89626479.43, '', ''),
    ]
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            put(wv, 3 + i, j, v, F_BOLD if i == 0 else (F_BLUE if isinstance(v, (int, float)) else F_BASE), NUM if isinstance(v, float) else None)
    style_header(wv, 3, 4)
    r = 3 + len(rows) + 1
    for n in ["Source: AP_Adv_payment_Oct_5th.xlsx, sheets 'Data W1', '4421', 'TCD (2)'; document-level comparison run on " + as_of + ".",
              "Conclusion: use sheet '4421'. It is five months more recent, carries the Nature (Avance / Amrt d'avance), Status and BU tags, agrees with the pivot the file was built on, and its JESA net balance ties to the AR PRE unamortised balances within 2.5% (see 'AR PRE check').",
              "'Data W1' is an earlier cut: it stops at April 2026 and still contains fully cleared advance / reversal pairs that the '4421' sheet removed.",
              "Not in scope of either version: GL 4421 for JETM (MA20), JEUS (US10), JEIN (MA30). The AR PRE shows ~1.5 bn MAD of unamortised client advances at JETM (OCP Green Water, water programme) that are missing from the AR side of this report."]:
        put(wv, r, 1, n, F_NOTE); r += 1
    widths(wv, [55, 30, 30, 26])

    wd = wb.create_sheet('README', 6)
    lines = [
        ('Weekly report: advances received from clients vs advances paid to suppliers', F_TITLE),
        ('', F_BASE),
        ('Definitions', F_BOLD),
        ('Gross ("excluding amortisation"): total advance invoiced to the client / paid to the supplier, before any deduction.', F_BASE),
        ('Amortised: the part of the advance already deducted from progress billing (client) or from supplier invoices (supplier).', F_BASE),
        ('Net ("including amortisation"): gross less amortised = balance still outstanding on the balance sheet (GL 4421 credit / GL 3411 debit).', F_BASE),
        ('Position AR - AP: net client advance held less net supplier advance paid. Positive = client cash funding exceeds what was passed to suppliers.', F_BASE),
        ('', F_BASE),
        ('Sources (this run)', F_BOLD),
        ("AR side: GL 4421000000 'Advance received from Customer' - sheet '4421' of AP_Adv_payment_Oct_5th.xlsx (JESA / MA10 only).", F_BASE),
        ("AP side: GL 3411000000 and 3411010000 advances to suppliers - sheet 'W1+W2 all items' of the same file (7 LEs, Dr = paid, Cr = amortised).", F_BASE),
        ("Check: AR_PRE_Week_3.xlsx sheet 'AR' - Status 1 = 'Advance' lines are the unamortised client advance per project (HT).", F_BASE),
        ('', F_BASE),
        ('Sheets', F_BOLD),
        ('Summary: LE x BU with AR gross / amortised / net, AP gross / amortised / net, and the position. AP category selector on Inputs!B4.', F_BASE),
        ('By LE: per legal entity, AP split Project vs Overhead/Annulation, tie-out of AP to the TB W2 balance in local currency.', F_BASE),
        ('By BU: per business unit across all LEs.', F_BASE),
        ('AR PRE check: GL 4421 vs AR PRE unamortised balance and gross advances, by LE x BU.', F_BASE),
        ('PRE vs GL by project: same check at project-root level, sorted by largest variance.', F_BASE),
        ('Version check: why sheet 4421 was retained over Data W1.', F_BASE),
        ('Inputs: as-of label, AP category selector, FX rates (yellow = update weekly), LE and BU mappings.', F_BASE),
        ('AR_GL / AP_GL / AR_PRE: line-item data with helper columns; every report figure is a SUMIFS on these sheets.', F_BASE),
        ('', F_BASE),
        ('Weekly refresh', F_BOLD),
        ('1. Get the new AP advances file (GL 3411 line items + GL 4421 line items, ideally for ALL company codes) and the new AR PRE file.', F_BASE),
        ('2. Run: python .claude/skills/advances-ar-vs-ap/scripts/build_advances_report.py --ap <AP file> --ar <AR PRE file> --out <report.xlsx> --as-of "Week N Mon YYYY"', F_BASE),
        ('3. Update FX rates on Inputs, check the control totals on Summary are zero and the TB difference on By LE is zero.', F_BASE),
        ('4. Review "AR PRE check" variances; investigate the top lines of "PRE vs GL by project".', F_BASE),
        ('', F_BASE),
        ('Colour code: blue = hard-coded input, black = formula, green = formula pulling from another sheet, yellow fill = update each week.', F_NOTE),
    ]
    for i, (t, f) in enumerate(lines, 1):
        c = wd.cell(row=i, column=1, value=t); c.font = f
    wd.column_dimensions['A'].width = 160

    for sh in wb.worksheets:
        sh.sheet_view.showGridLines = False if sh.title in ('Summary', 'By LE', 'By BU', 'AR PRE check', 'README', 'Version check', 'Inputs') else True
    wb.save(out)
    return {'ar_rows': len(argl), 'ap_rows': len(apgl), 'pre_rows': len(pre)}


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--ap', required=True); p.add_argument('--ar', required=True); p.add_argument('--out', required=True)
    p.add_argument('--as-of', default='Week 3 Oct 2026'); p.add_argument('--ar-sheet', default='4421')
    p.add_argument('--ap-sheet', default='W1+W2 all items'); p.add_argument('--pre-sheet', default='AR')
    p.add_argument('--usd', type=float, default=9.0); p.add_argument('--xof', type=float, default=0.0165)
    a = p.parse_args()
    print(build(a.ap, a.ar, a.out, a.as_of, a.ar_sheet, a.ap_sheet, a.pre_sheet, a.usd, a.xof))
