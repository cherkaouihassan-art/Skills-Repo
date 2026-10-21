# Weekly AR, Cash & Payables deck — JESA

A repeatable pipeline that turns four weekly source files into the executive PowerPoint
("AR, Cash & Payables Position – Weekly Executive Review").

```
sources (4 files)  ──►  extract_inputs.py  ──►  inputs/weekly_inputs.xlsx  ──►  build_deck.py  ──►  output/*.pptx
                                                 (review / edit by hand)
```

## Weekly run (about 5 minutes)

1. Drop the four refreshed files anywhere (names can change, paths are parameters):
   * **AR dashboard** – `AR_PRE.xlsx` (sheets `Executive Summary` + `Comparatives`). Before extracting, make sure the
     dashboard's *Snapshot* tab holds last week and the *AR Baseline* tab holds the month-start position.
   * **Cash flow workbook** – `CASH_FLOW_ANALYSIS - BUs <month> W<n>.xlsx` (sheets `EXEC Summary`, `Approved vs Tracking log …`,
     `Priority list …`, `LC Ongoing …`).
   * **Bank LC situation** – `LCs SITUATION - <date>.xlsx` (sheet `LC avec echéance`).
   * **Supplier invoices** – `Supplier Invoices - <month> W<n>.xlsx` (sheet `4411 - Parked - Downpayment`).
   * **UR Tracking Consolidation** – unrecognized revenue (`Summary UR`, `Sheet1` previous BU update, `Summary PO by BU`).
   * **Beginning-of-month AR file** – the raw AR line list saved at the start of the month (`AR_Beg_of_the_month.xlsx`,
     sheet `AR`). Same file all month; it is the "vs month start" baseline.
2. Extract the inputs (dates are the only things to type):
   ```bash
   python3 extract_inputs.py \
     --ar AR_PRE.xlsx --cash "CASH_FLOW_ANALYSIS - BUs October W3.xlsx" \
     --lc "LCs SITUATION - October 19th.xlsx" --ap "Supplier Invoices - Oct W3.xlsx" \
     --ur "UR Tracking Consolidation.xlsx" \
     --ar-baseline "AR_Beg_of_the_month.xlsx" --ar-baseline-date 2026-10-07 \
     --as-of 2026-10-21 --ar-date 2026-10-14 --ar-prev-date 2026-10-07 \
     --collections-to 19/10 --payments-to 15/10 --lc-date 19/10 --ap-label "Oct W3" \
     --out inputs/weekly_inputs.xlsx
   ```
   Without `--ar-baseline` the month-start columns fall back to the dashboard's own *AR Baseline* tab; without `--ur`
   the unrecognized revenue comes from the dashboard.
3. Open `inputs/weekly_inputs.xlsx` and review. Everything on the slides comes from this workbook:
   * `KEY_MESSAGES` – five auto-drafted messages; rewrite the wording freely.
   * `COLLECTIONS_MONTHLY` – labor / non-labor split per month (the current month total is refreshed from the AR
     dashboard; overwrite the split if you have the exact figures).
   * any number can be corrected by hand before building.
4. Build the deck:
   ```bash
   python3 build_deck.py --inputs inputs/weekly_inputs.xlsx --out output/JESA_AR_Cash_Position_2026-10-21.pptx
   ```

## Input workbook structure (`inputs/weekly_inputs.xlsx`)

| Tab | Content (MMAD) | Feeds slides |
|---|---|---|
| `META` | as-of dates, AR snapshot / previous / baseline dates, cut-off labels | all |
| `KEY_MESSAGES` | n, title, body | 2 |
| `AR_KPI` | metric × month_start / prev_week / current (advance, unbilled, unpaid, total AR, unrecognized, gross exposure, AP buckets, LC) | 1, 3, 4, 12 |
| `AR_SPLITS` | labor / non-labor and OCP / non-OCP (current) | 3 |
| `AR_BU` | BU × component × month_start / prev_week / current | 4, 5 |
| `AR_UNBILLED_OWNER` | BU × (at risk / under JESA / under OCP / SS postponed) | 6 |
| `AR_UNBILLED_STATUS` | detailed unbilled statuses (OCP / JESA / at risk) | 6 |
| `AR_UNPAID_STATUS` | unpaid by OCP processing status | 6 |
| `AR_AGING` | AR by year of origin | 4 |
| `UNRECOGNIZED` | unrecognized suppliers by BU and status | 7 |
| `COLLECTIONS_BU`, `COLLECTIONS_MONTHLY` | payments received by BU (month to date, prior month, YTD) and monthly series | 8 |
| `CASH_MONTHLY` | group cash IN / OUT / net by month incl. October commitments, PSR block | 9 |
| `CASH_BU`, `CASH_BU_OCT` | cumulative and current-month cash position by BU | 10, 11 |
| `ONGOING` | paid / approved (LC, suppliers) / pending, tracking-log status, priority list | 9, 11 |
| `AP_BU` | payables & LC by BU, previous vs current week (AR dashboard) | 12 |
| `AP_BUCKETS`, `AP_BU_ORIGIN`, `AP_TIERS`, `AP_TOP_SUPPLIERS` | supplier-invoice file analysis | 12, 15 |
| `LC_LINES` | one line per LC with maturity schedule | 13, 14 |

## Comparatives

* **Week on week** = AR dashboard *Snapshot* tab (previous week) vs current.
* **vs month start** = beginning-of-month AR line list aggregated with the dashboard rules (unbilled excludes
  `J-6-a Unearned on LS`; advance excludes `J-6-*` and `C-3-a` lines; at risk = `C-1-a/b`, SS postponed = `J-3-c`,
  under JESA = other `J-*`, under OCP = `C-*` and PDR). The aggregation reproduces the dashboard snapshot to within 2 MMAD.
* **Unrecognized revenue WoW** = `Summary UR` grand total (incl. markup) vs the "Previous BU Update" block including
  markup 6% in `Sheet1` of the UR consolidation.

## Conventions

* All amounts in MMAD (millions of MAD); AR decreases are shown green, increases red; cash net balances the other way round.
* Business-unit names are normalised (`AMT`→AM, `FEC`→Fertilizers & Chemicals, `WEN`→Water & Environment, …) in `extract_inputs.py`.
* The deck uses a text "JESA" wordmark top-right; drop the logo image on the slide master if preferred.

Dependencies: `python3`, `openpyxl`, `python-pptx` (`pip install openpyxl python-pptx`).
