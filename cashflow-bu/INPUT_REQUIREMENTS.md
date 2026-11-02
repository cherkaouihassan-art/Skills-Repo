# CASH FLOW ANALYSIS – BUs — required inputs for a refresh

This is the input structure for the daily/weekly refresh of `CASH FLOW ANALYSIS - BUs`. Each refresh takes up to **seven input streams**. Send what has moved; the rest is carried forward from the last load. Every stream has a native file we can read directly, a minimum column set, and a tab in the intake pack (`templates/CF_INPUT_PACK.xlsx`) for the cases where the data only exists in an email, a screenshot or a PDF.

Current state of each stream is in `state/refresh-state.json`; the human-readable history is `state/refresh-log.md`.

## Stream summary

| # | Stream | Feeds | Native file (preferred) | Owner / source report | Frequency | As-of at W3 |
|---|---|---|---|---|---|---|
| 1 | Cash IN – OCP collections | Report cash-in rows (PSR / Procurement by BU) | `OCP Payments <month>.xlsx` — JESA OCP File nn (AX export) | Weekly OCP payments report | every refresh | OCP File 26, paid to 19/10 |
| 2 | Cash OUT – SAP payments cleared | October actual column (E) | `Suppliers Payments by BU <date>.xlsx` / `PAYMENTS OCT …` (FBL1N-style extract) | Treasury / AP | every refresh | cleared to 15/10 (W3 file 22.10) |
| 3 | Approved payments | Ongoing position, col F | Payment tracking log – Invoice Tracker export, your approval emails, approved lists per BU | You (approvals) + AP tracking log | every refresh | tracking log W3 (20/10) |
| 4 | Pending approval | Ongoing position, col H | Lists sent to you for approval (BU batches, invoice lists, PDFs) | BUs / AP | as received | WEN Batch 5 (19/10), AM W1 (12/10) |
| 5 | LC payments and LC situation | LC rows (actual + ongoing) | `<date> - LC PAYMENTS.xlsx` (bank lines) + `LCs SITUATION - <date>.xlsx` sheet `LC avec echéance` | Treasury | weekly | LC paid 21/10, situation 19/10 v003 |
| 6 | Customs duties | Customs rows | Treasury customs detail (TGR / RECEVEUR DOUANE) by BU and date | Treasury | weekly | to 15/10 (tab `DD`) |
| 7 | Standing inputs | PSR reserve, dividends, P&L revenue, master switch, as-of date | Your instruction / monthly P&L | You / Finance | monthly or on change | Sept revenue loaded 06/10 |

## What to send, per stream

### 1. Cash IN – OCP payment report
- **Send:** the latest OCP payments workbook as-is (`OCP Payments <month>.xlsx`), or the AX export sheet. It is cumulative; we load only new invoices.
- **Minimum columns:** `Réference` (invoice), `Fournisseur`, `Net à payer`, `Date Paiement`, `Réf Paiement`, `Pièce comptable`, `Division`, `Commande`, `Filiale`. Plus the `Summary by BU` tab if the file has it — we tie to it.
- **Also useful:** the Power BI - AR extract when new projects appear (gives PSR/PROC and BU for unmapped invoices).
- **Intake tab if no file:** `1_CashIn_OCP`.
- **We confirm back:** new invoices loaded, October total by BU vs OCP summary, any `Unmapped` invoices needing a BU.

### 2. Cash OUT – SAP payments cleared
- **Send:** the SAP payment extract for the period (one file per cut: 1–8/10, 1–15/10, 1–22/10 …). Same layout as `October Actual Payments` columns A:AE.
- **Minimum columns:** `Type`, `Montant en MAD`, `Paiement positive`, `Supplier`, `Name`, `Document Number`, `Document Type`, `Reference`, `Clearing Document`, `Clearing Date`, `Posting Date`, `Purchasing Document`, and `BU` if the extract carries it.
- **Intake tab if no file:** `2_CashOut_SAP`.
- **We confirm back:** payments added, ongoing lines closed (reference match), short-paid balances created, payments not on any ongoing list that need a BU from you.

### 3. Approved payments (your approvals)
Three accepted forms, any mix:
- **a. Payment tracking log** – export of the `Invoice Tracker` table (`Tracking.xlsx` / `Payment tracking log Wn.xlsx`). Columns: `Request date`, `BU`, `Supplier Name`, `Invoice Reference`, `Invoice Amount TTC`, `Currency`, `PO`, `Accounting Status`, `Type de blocage`, `Treasury Comments`, `Payment Date`, `Overall status`. Give us read access to the live file and we re-export it ourselves at each refresh.
- **b. Approval email** – forward it or paste it. We need, per line: BU, supplier, invoice/reference, amount TTC MAD (or FX + currency), and the date of your approval. Attachments (invoice PDFs) are read for the reference and amount.
- **c. Approved list file** – any BU list you approved (priority list, AM approved list, WEN batch, ENE list). Minimum: BU, supplier, PO, reference, amount, your approval date.
- **Intake tab if no file:** `3_Approved`.
- **Rule:** a line is "approved" when it is in the tracking log or on a list you approved. A line the log says is PAID stays open until SAP shows the payment.

### 4. Pending approval (sent to you, not yet approved)
- **Send:** the lists/invoices that reached you for approval (BU batches, AP lists, scanned invoices). Minimum per line: BU, supplier, PO, reference, amount TTC MAD, due date, tier if known, source document.
- **Intake tab:** `4_Pending` (same layout as `3_Approved`, plus `Due date`).
- When you later approve some of them, just tell us which (reference list or email) — they move from pending (col H) to approved (col F) without re-keying.

### 5. LC payments and LC situation
- **LC paid (actual):** Treasury's `PAIEMENTS PAR LC` bank lines: `Date`, `Libellé`, `Débit (DH)`, `MNT DEVISE`, `D.Curr`, `PO`, `Ref LC`, `Nom fournisseur`, `Project` code (AMT/FEC/WEN/SP2M/ENE/JET).
- **LC ongoing:** the bank LC situation file, sheet `LC avec echéance`: `L.E.`, `Project code`, `PO`, `LC / CREDOC ref`, `Supplier`, `LC expiry`, `Ccy`, `LC face value`, `Encours LC (FX)`, `Encours LC (MAD)`, `<month> maturity (MAD)`, `Action`, `Comments`.
- **Intake tabs:** `5a_LC_Paid`, `5b_LC_Situation`.
- **We confirm back:** LCs paid in the month that drop out of ongoing, maturities by BU.

### 6. Customs duties
- **Send:** Treasury customs detail (what fed tab `DD`): `Date`, `Valeur`, `Libellé d'opération`, `Montant sur relevé`, `Montant du ticket`, `BU` code. One line per BU per TGR payment.
- **Intake tab:** `6_Customs`.

### 7. Standing inputs (only when they change)
- Master switch (include October ongoing in cumulative: TRUE/FALSE) and the data-as-of date to print on the cover.
- PSR reserve total and months; PROC compensable (dividend) schedule by BU; Global Mining dividend invoices.
- Monthly P&L revenue (PSR and PROC, MMAD, group and per BU) when a month closes.
- Month roll-forward instruction (when November actuals start: new SAP tab + new report column).
- **Intake tab:** `7_Standing`.


## Where the inputs live today (Google Drive, connected)
| Stream | Drive location | File |
|---|---|---|
| 1 Cash IN | `AR CASH DECK  OCP Payments Weekly/` | `OCP Payments Weekly.xlsx` — output of the weekly OCP payment report conversation; sheet `Payments Key` is per invoice with BU, Class (Procurement = PROC, Proservices = PSR) and Pay Month, `Unmapped Invoices` lists what still needs a BU |
| 2 SAP payments | `COO Report/Input/` | `PAYMENTS OCTOBER W<n> _ dd.mm.yy.xlsx`, sheet `PAYMENTS OCT` |
| 3 Tracking log | `COO Report/Input/` | `Payment tracking log.xlsx`, sheet `Invoice Tracker` (header on row 3) |
| 5b LC situation | `AR CASH DECK/Inbox/` | `LCs SITUATION - <date>.xlsx` |
| 5a LC paid, 6 customs, 4 pending lists | not on Drive yet | drop them in `AR CASH DECK/Inbox/` or upload in chat |
File IDs and what was last seen in each file are kept in `state/refresh-state.json → drive_sources`. The refresh takes the newest file in each folder; replacing the file in place (same name) is fine.

## Keys we match on (so lines are never double counted)
| Stream | Dedupe key | Match to ongoing position |
|---|---|---|
| Cash IN | invoice ref + payment ref + amount | n/a |
| SAP payments | document no. + clearing doc + amount | normalised invoice reference → then supplier + amount (flagged WEAK MATCH) |
| Approved / pending | normalised reference + supplier | same |
| LC paid | date + amount + LC ref | LC situation line (sets ongoing to 0) |
| Customs | date + amount + BU code | n/a |

## How a refresh runs (what you get back)
1. You drop files in `cashflow-bu/inbox/<yyyy-mm-dd>/` (or upload them in chat, or forward emails) and say "update the cash flow report".
2. We normalise each file into the stream above, load it into the matching source tab, log a `Data Sources` row per load and update `state/refresh-state.json`.
3. Full recalculation, 18 integrity checks must read OK, headline reconciled stream by stream to the previous version.
4. You receive the refreshed workbook (`CASH FLOW ANALYSIS - BUs <Month> W<n>.xlsx`), a one-page movement summary (previous → new headline by stream), and the open questions list (unmapped BUs, untiered suppliers, log-paid-not-in-SAP lines).

## Open questions list format (what we will ask you to rule on)
- Payments in SAP not on any ongoing list → BU?
- New suppliers with no tier → tier?
- Lines the tracking log says PAID with no SAP payment → chase or close?
- Pending lines older than 30 days → still pending, or deferred?
