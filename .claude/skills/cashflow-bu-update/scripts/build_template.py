#!/usr/bin/env python3
"""Build the intake pack template CF_INPUT_PACK.xlsx (one tab per input stream)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

out = sys.argv[1] if len(sys.argv) > 1 else 'CF_INPUT_PACK.xlsx'
BUS = ['AM', 'Advisory', 'Building & Capabilities', 'Building & Infrastructure', 'Energy Operations', 'Fertilizers & Chemicals',
       'Global Mining', 'JESA Technologies', 'SP2M Ad Hoc Program', 'Water & Environement']
TIERS = ['1', '2', '3', 'Intl 1', 'Intl 2', 'Intl 3', 'Etranger*', 'OCP et Filiale', 'Office/Publique', 'Unclassified', 'Not on OCP Tier']
F = 'Arial'
hdr_fill = PatternFill('solid', start_color='1F3864'); hdr_font = Font(name=F, bold=True, color='FFFFFF', size=10)
req_fill = PatternFill('solid', start_color='FFF2CC'); note_font = Font(name=F, italic=True, size=9, color='595959')
title_font = Font(name=F, bold=True, size=13); input_font = Font(name=F, color='0000FF', size=10)
thin = Side(style='thin', color='BFBFBF')

TABS = [
 ('README', None, None, None),
 ('1_CashIn_OCP', 'Cash IN — OCP payment report (one row per paid invoice). Preferred: send the OCP Payments workbook itself; use this tab only for lines that exist nowhere else.',
  [('Réference (invoice)*', 16), ('Fournisseur*', 28), ('Net à payer (MAD)*', 16), ('Date Paiement*', 13), ('Réf Paiement', 14), ('Pièce comptable', 14),
   ('BU*', 24), ('Labor / Non-Labor*', 14), ('Division', 12), ('Commande (PO)', 14), ('Filiale', 10), ('ID DOC', 12), ('Source document / note', 40)],
  {'BU*': 'BU', 'Labor / Non-Labor*': 'LAB'}),
 ('2_CashOut_SAP', 'Cash OUT actuals — SAP payment lines cleared (one row per payment document line). Amount as in SAP (positive for a payment); only Paiement lines with Paiement positive = Oui are counted.',
  [('Type*', 10), ('Montant en MAD*', 16), ('Paiement positive*', 10), ('Supplier (code)', 12), ('Name*', 30), ('Document Number*', 14), ('Document Type', 8),
   ('Reference (invoice)*', 18), ('Clearing Document*', 14), ('Clearing Date*', 13), ('Posting Date', 13), ('Purchasing Document (PO)', 16), ('BU (if known)', 24),
   ('Tier (if known)', 14), ('Source document / note', 40)],
  {'BU (if known)': 'BU', 'Tier (if known)': 'TIER', 'Paiement positive*': 'OUI'}),
 ('3_Approved', 'Approved payments — lines you approved (email, approved list, tracking log). Amount TTC in MAD, cash out NEGATIVE. One row per invoice / attachment / balance.',
  [('BU*', 24), ('Supplier*', 30), ('PO', 16), ('Reference / Invoice*', 20), ('Amount (MAD, negative)*', 18), ('Currency', 8), ('FX amount', 14), ('Due date', 13),
   ('Tier', 14), ('Approval date*', 13), ('Approved via* (email / list / tracking log)', 22), ('In tracking log? (Yes/No)', 12), ('Priority / Status', 20), ('Comment', 30), ('Source document*', 40)],
  {'BU*': 'BU', 'Tier': 'TIER', 'In tracking log? (Yes/No)': 'YN'}),
 ('4_Pending', 'Pending approval — lines sent to you and not yet approved. Same layout as 3_Approved; when you approve them later, list the references and they move to approved.',
  [('BU*', 24), ('Supplier*', 30), ('PO', 16), ('Reference / Invoice*', 20), ('Amount (MAD, negative)*', 18), ('Currency', 8), ('FX amount', 14), ('Due date*', 13),
   ('Tier', 14), ('Date received*', 13), ('Requested by (BU / AP)', 20), ('In tracking log? (Yes/No)', 12), ('Priority / Status', 20), ('Comment', 30), ('Source document*', 40)],
  {'BU*': 'BU', 'Tier': 'TIER', 'In tracking log? (Yes/No)': 'YN'}),
 ('5a_LC_Paid', 'LC payments — bank statement LC lines (Treasury "PAIEMENTS PAR LC"). Débit positive as on the statement.',
  [('Date*', 13), ('Libellé', 30), ('Débit (DH)*', 16), ('MNT DEVISE', 14), ('D.Curr', 8), ('PO', 14), ('Ref LC / CREDOC*', 20), ('Nom fournisseur*', 30),
   ('Project code* (AMT/FEC/WEN/SP2M/ENE/JET)', 16), ('Tier', 14), ('Source document*', 40)],
  {'Project code* (AMT/FEC/WEN/SP2M/ENE/JET)': 'PROJ', 'Tier': 'TIER'}),
 ('5b_LC_Situation', 'LC situation — bank "LC avec echéance" sheet: every open LC with the maturity expected in the current month. Positive amounts; the loader signs them.',
  [('Source row', 8), ('L.E.', 8), ('Project code*', 12), ('PO', 14), ('LC / CREDOC ref*', 20), ('Supplier*', 30), ('LC expiry', 13), ('Ccy', 6), ('LC face value (FX)', 16),
   ('Encours LC (FX)', 16), ('Encours LC (MAD)', 16), ('Current-month maturity (MAD)*', 18), ('Action', 20), ('Comments', 40), ('Business Unit*', 24), ('Already paid this month? (Yes/No)', 12), ('Situation date*', 13)],
  {'Project code*': 'PROJ', 'Business Unit*': 'BU', 'Already paid this month? (Yes/No)': 'YN'}),
 ('6_Customs', 'Customs duties — Treasury customs detail (TGR / RECEVEUR DOUANE) by BU and date. One row per BU per payment, amount positive as on the statement.',
  [('Company code', 10), ('Date*', 13), ('Valeur', 13), ("Libellé d'opération*", 34), ('Montant sur relevé (MAD)*', 18), ('Montant du ticket', 16), ('BU code* (AMT/FEC/WEN/SP2M/ENE/GMM/BIN/JET/ADS/BIC)', 16), ('Source document*', 40)],
  {'BU code* (AMT/FEC/WEN/SP2M/ENE/GMM/BIN/JET/ADS/BIC)': 'CODE'}),
 ('7_Standing', 'Standing inputs — fill only the lines that change.',
  [('Input', 44), ('Value', 20), ('Effective date', 13), ('Instruction / source', 50)], {}),
]

wb = Workbook(); wb.remove(wb.active)
lists = wb.create_sheet('Lists')
for i, b in enumerate(BUS, 1): lists.cell(i, 1, b)
for i, t in enumerate(TIERS, 1): lists.cell(i, 2, t)
for i, c in enumerate(['AMT', 'ADS', 'BIC', 'BIN', 'ENE', 'FEC', 'GMM', 'JET', 'SP2M', 'WEN'], 1): lists.cell(i, 3, c)
for i, c in enumerate(['AMT', 'ENE', 'FEC', 'JET', 'SP2M', 'WEN'], 1): lists.cell(i, 4, c)
for i, c in enumerate(['Labor', 'Non-Labor'], 1): lists.cell(i, 5, c)
for i, c in enumerate(['Yes', 'No'], 1): lists.cell(i, 6, c)
for i, c in enumerate(['Oui', 'Non'], 1): lists.cell(i, 7, c)
lists.sheet_state = 'hidden'
DV = {'BU': "=Lists!$A$1:$A$10", 'TIER': "=Lists!$B$1:$B$11", 'CODE': "=Lists!$C$1:$C$10", 'PROJ': "=Lists!$D$1:$D$6",
      'LAB': "=Lists!$E$1:$E$2", 'YN': "=Lists!$F$1:$F$2", 'OUI': "=Lists!$G$1:$G$2"}

for name, desc, cols, dvmap in TABS:
    ws = wb.create_sheet(name, index=len(wb.sheetnames) - 1)
    if name == 'README':
        ws.column_dimensions['A'].width = 4; ws.column_dimensions['B'].width = 120
        ws['B2'] = 'CF INPUT PACK — CASH FLOW ANALYSIS - BUs'; ws['B2'].font = title_font
        lines = [
         'Use this pack when an input arrives without a native file (email approvals, screenshots, PDFs, verbal instructions).',
         'Native files (OCP Payments workbook, SAP payment extract, tracking log export, LC files, customs detail) are sent as-is — do not retype them here.',
         '',
         'Rules: one row per line · amounts in MAD · cash OUT negative on tabs 3 and 4, as-on-statement elsewhere · dates as dd/mm/yyyy or Excel dates · columns marked * are required.',
         'BU names, tiers and codes are drop-downs (exact spelling matters: "Water & Environement" with a single n is the workbook convention).',
         'Keep the original reference exactly as on the document; the loader normalises it for matching.',
         'Always fill "Source document": file name, email subject + date, or screenshot name. It is copied to the Data Sources audit trail.',
         '',
         'Tabs: 1_CashIn_OCP · 2_CashOut_SAP · 3_Approved · 4_Pending · 5a_LC_Paid · 5b_LC_Situation · 6_Customs · 7_Standing',
         'Save as CF_INPUT_PACK_<yyyy-mm-dd>.xlsx and drop it in cashflow-bu/inbox/<yyyy-mm-dd>/ with the native files.',
        ]
        for i, l in enumerate(lines, 4): ws.cell(i, 2, l).font = Font(name=F, size=10)
        continue
    ws['A1'] = name.replace('_', ' ') + ' — ' + desc; ws['A1'].font = Font(name=F, bold=True, size=11); ws['A1'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(6, len(cols))); ws.row_dimensions[1].height = 42
    ws['A2'] = 'Columns marked * are required. Rows start at line 4. Blue = your input.'; ws['A2'].font = note_font
    for j, (h, w) in enumerate(cols, 1):
        c = ws.cell(3, j, h); c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(wrap_text=True, vertical='center')
        c.border = Border(bottom=thin); ws.column_dimensions[get_column_letter(j)].width = w
        if h.endswith('*') or '*' in h:
            for r in range(4, 204): ws.cell(r, j).fill = req_fill
        for r in range(4, 204): ws.cell(r, j).font = input_font
        if h in dvmap:
            dv = DataValidation(type='list', formula1=DV[dvmap[h]], allow_blank=True); ws.add_data_validation(dv); dv.add(f'{get_column_letter(j)}4:{get_column_letter(j)}203')
        if 'Date' in h or 'date' in h or 'expiry' in h:
            for r in range(4, 204): ws.cell(r, j).number_format = 'dd/mm/yyyy'
        if 'MAD' in h or 'Montant' in h or 'Débit' in h or 'amount' in h.lower() or 'Encours' in h or 'value' in h.lower():
            for r in range(4, 204): ws.cell(r, j).number_format = '#,##0.00;(#,##0.00);-'
    ws.row_dimensions[3].height = 44; ws.freeze_panes = 'A4'
    if name == '7_Standing':
        for i, (k, note) in enumerate([
            ('Master switch — include October ongoing in cumulative (TRUE/FALSE)', "EXEC Summary!E6"),
            ('Data as of (cover date)', "EXEC Summary!E7"),
            ('PSR reserve — total (MAD)', "PSR Reserve (provision)!B5"),
            ('PSR reserve — number of months', "PSR Reserve (provision)!B6"),
            ('PROC compensable (dividends) — AM', "PROC Compensable (standing)!C5"),
            ('PROC compensable — Advisory', 'C6'), ('PROC compensable — Building & Capabilities', 'C7'), ('PROC compensable — Building & Infrastructure', 'C8'),
            ('PROC compensable — Energy Operations', 'C9'), ('PROC compensable — Fertilizers & Chemicals', 'C10'), ('PROC compensable — Global Mining', 'C11 (formula: base + GM dividend invoices)'),
            ('PROC compensable — JESA Technologies', 'C12'), ('PROC compensable — SP2M Ad Hoc Program', 'C13'), ('PROC compensable — Water & Environement', 'C14'),
            ('Monthly revenue PSR — JESA Group (MMAD) for month …', "SRC - Monthly Revenue (<month>)!F7"),
            ('Monthly revenue PROC — JESA Group (MMAD) for month …', "SRC - Monthly Revenue (<month>)!F8"),
            ('Month roll-forward requested (yes + month)', 'new SAP tab + new report column'),
        ], 4):
            ws.cell(i, 1, k).font = Font(name=F, size=10); ws.cell(i, 4, note).font = note_font
wb.save(out); print('written', out)
