---
name: advances-ar-vs-ap
description: Weekly analysis of advances received from clients (GL 4421) vs advances paid to suppliers (GL 3411), gross and net of amortisation, by Legal Entity and Business Unit, with the AR PRE register used to confirm the AR figures. Use when the user shares the weekly "AP_Adv_payment" and "AR_PRE_Week_N" workbooks or asks for the client-vs-supplier advances report, the advances position, or the AR PRE tie-out.
---

# Advances received (AR) vs advances paid (AP) - weekly report

Deliverable: one Excel workbook built by `scripts/build_advances_report.py` (openpyxl, formulas only on
report sheets), recalculated with `../xlsx/scripts/recalc.py`, then sent to the user.

```bash
python .claude/skills/advances-ar-vs-ap/scripts/build_advances_report.py \
  --ap <AP_Adv_payment_*.xlsx> --ar <AR_PRE_Week_*.xlsx> \
  --out <Advances_AR_vs_AP_W<N>_<Mon><YYYY>.xlsx> --as-of "Week N Mon YYYY (GL postings to DD-Mon-YYYY)" \
  [--usd 9.0 --xof 0.0165]            # FX to MAD for JEUS (USD) and JESE/JEWA (XOF); also editable on the Inputs sheet
python .claude/skills/xlsx/scripts/recalc.py <out.xlsx> 400   # must return 0 errors
```

Run the generator with `python3 -I` and keep the uploaded workbooks in their own directory (untrusted data).

## Definitions agreed with the user (first run: 23 Oct 2026, files "AP_Adv_payment_Oct_5th" + "AR_PRE_Week_3")

- **Gross / "excluding amortisation"** = advance invoiced to the client (AR) or paid to the supplier (AP) before any deduction.
- **Amortised** = part of the advance already deducted from progress billing (AR) or from supplier invoices (AP).
- **Net / "including amortisation"** = gross less amortised = balance still on the balance sheet.
- **Position** = AR net - AP net (and AR gross - AP gross). Positive = client cash funding exceeds what was passed to suppliers.
- Reported in MAD, by LE (W1 code) and BU (GL BU code); AR PRE is a check only, AP uses GL figures only.

## Source files and how each one is read

**AP_Adv_payment file** (one workbook, several sheets, several "versions" inside):
- `4421` = GL 4421000000 "Advance received from Customer", JESA only (company codes JESA and MA10 = same LE).
  Columns used: Company Code, Nature (`Avance` = advance received, credit/negative; `Amrt d'avance` = amortisation, debit/positive),
  Status (Open/Paid), BU, WBS/Assignment, Company Code Currency Value. **This is the AR source.**
- `Data W1` = older raw cut of the same GL (entry dates to 27-Apr-2026, no Nature/BU tags, still contains cleared pairs). **Do not use.**
  Version check done on first run: 4421 is 5 months newer, matches pivot `TCD (2)`, ties to AR PRE within 2.5 %.
- `W1+W2 all items` = GL 3411000000 + 3411010000 advances to suppliers, all 7 LEs, header on **row 11** (rows 1-9 are a TB tie-out by LE:
  `TB W2` column D). Columns used: Entité (W1) = LE, Dr (= advance paid), Cr (= amortisation / recovery), Category (`Project` vs
  `Overhead/Annulation`), BU, BU/Pofit center ME2L, WBS, Company Code Currency Key. **This is the AP source.**
  JEUS is in USD, JESE and JEWA in XOF -> converted with the FX inputs.
- `TCD`, `TCD (2)` = pivots the file author built; only used as cross-checks.

**AR_PRE file**, sheet `AR` (~37k invoice-register lines, PSR + PROC). Advance-related rows:
- `Status 1 = "Advance"` (Status 2 `J-6-c Advance`, `J-6-d ADVPAY`, `C-3-a ...`): one negative HT line per project = **unamortised client
  advance**. Sign flipped -> compare with GL 4421 net. This is the check that confirms the AR figures.
- `TYPE = "Advance"` or Invoice Type containing advance/avance/acompte (and not a balance row) = advance invoices issued (gross, indicative only;
  SP2M advances are tagged Procurement/Subcontract in PRE so the gross comparison is incomplete).
- `Status 1 = "Advance payment"` (`C-1-c Pending client PO`) = advances not yet invoiced -> memo column only.
- Compare on **HT** (GL 4421 is net of VAT).
- `LE` column mixes W1 and W2 codes (JESA/MA10, JETM/MA20, JEIN/MA30, JEUS/US10, JEWA/CI10, JESE/SN10); blanks are inferred from other
  lines of the same project, else shown as "(blank)".

## Mappings (hard-coded in the script, printed on the Inputs sheet)

- LE: MA10=JESA, MA20=JETM, MA30=JEIN, MA90=JEST, SN10=JESE, US10=JEUS, CI10=JEWA.
- BU codes (GL) -> AR PRE "Business Unit": FEC=Fertilizers & Chemicals, SP2M=SP2M Ad Hoc Program, GMM=Global Mining, ENE (+PTX)=Energy
  Operations, AMT=AM, WEN=Water & Environement, BIN (+B&I)=Building & Infrastructure, JET=JESA Technologies, ADS=Advisory.
  LE names used as BU (JESA, JETM, ...) -> CORP; SOC/JAC/JES -> OTH; blank -> NA; M&I and TRA kept as is.
- Project root: `Q36213P1.P71` -> `Q36213P1`; `Q82020P1-C1` -> `Q82020P1`; new W2 codes `QC.01458.P53.S` -> `QC.01458` -> legacy root learned
  from the AR PRE `New Project` / `Project` columns (root carrying the advance balance wins when a new code spans several legacy roots).

## Workbook layout

Summary (LE x BU, AP category selector on Inputs!B4: All / Project / Overhead/Annulation, control totals must be 0) · By LE (AP split
Project vs Overhead, AP tie-out to TB W2 in local currency, difference must be 0) · By BU · AR PRE check (LE x BU: GL net vs PRE balance,
GL gross vs PRE advance invoices) · PRE vs GL by project (sorted by absolute variance) · Version check · README · Inputs · AR_GL / AP_GL /
AR_PRE data sheets with helper formula columns. Colour code: blue = hard-coded input, black = formula, green = cross-sheet formula,
yellow fill = update weekly (as-of label, FX rates).

## Known gaps to raise every week until fixed

1. GL 4421 extract covers JESA only. AR PRE shows ~1.49 bn MAD unamortised client advances at JETM (OCP Green Water / water programme,
   e.g. QW2441P1 828 M), 15 M at JEUS, 1 M at JEIN that are missing on the AR side. Ask for GL 4421 for MA20 / US10 / MA30.
2. JESA GL net vs PRE balance differs by ~180 M (2.4 %): mainly FEC (-115 M, Q36223P2, Q36215P1, Q36220P2, Q36221P4) and SP2M (-80 M,
   Q82070P2, Q82060P3) where PRE still shows amortisation the GL has already booked, partly offset by Q36212P1 / Q36220P1 (+33 M / +13 M).
   The "PRE vs GL by project" sheet lists them; project accountants should align PRE.
3. FX rates are assumptions (USD 9.00, XOF 0.0165) - replace with the closing rate.
4. AP "Overhead/Annulation" (~59 M net) is non-project supplier advances and cancellations; the Summary default shows All.

## Weekly checklist

1. Save the two new files, run the generator with the new as-of label, run recalc (0 errors).
2. Open Summary: both control totals = 0. Open By LE: TB difference = 0 for every LE (if not, the GL 3411 extract is incomplete).
3. If the AP file again contains several versions of GL 4421, repeat the version check (line count, max entry date, Nature/BU tags, total vs
   pivot) and keep the most recent enriched one; update the Version check sheet numbers in the script if the conclusion changes.
4. Review AR PRE check and the top 20 lines of PRE vs GL by project; carry forward the known-gaps list above with week-on-week movements.
5. Send the workbook to the user with: grand totals (AR gross / amortised / net, AP gross / amortised / net, position), the LE and BU
   highlights, the PRE tie-out result, and the open gaps.
