# Weekly AR, Cash & Payables deck — JESA

A repeatable pipeline that turns four weekly source files into the executive PowerPoint
("AR, Cash & Payables Position – Weekly Executive Review").

```
sources (4 files)  ──►  extract_inputs.py  ──►  inputs/weekly_inputs.xlsx  ──►  build_deck.py  ──►  output/*.pptx
                                                 (review / edit by hand)
```

## Weekly run (about 5 minutes)

1. Drop the refreshed files anywhere (names can change, paths are parameters):
   * **Raw AR extract of the week** – `AR_PRE_Week_<n>.xlsx`, single sheet `AR` (one line per invoice / unbilled line).
     `ar_engine.py` aggregates it with the dashboard rules; the totals reproduce the BU meeting Power BI report exactly
     (e.g. week 3: total 12,395.1 MMAD, unbilled 9,477.5, unpaid 2,369.6, advance 548.0).
   * **Previous week** – last week's `inputs/history/weekly_inputs_<date>.xlsx` (`--prev-inputs`), or last week's raw AR
     extract (`--ar-prev`, gives per-BU status deltas as well).
   * **Beginning-of-month AR file** – the raw AR line list saved at the start of the month (`--ar-baseline`). Same file all
     month; it is the "vs month start" baseline.
   * **Cash flow workbook** – `CASH_FLOW_ANALYSIS - BUs <month> W<n>.xlsx` (sheets `EXEC Summary`, `Approved vs Tracking log …`,
     `Priority list …`, `LC Ongoing …`).
   * **Bank LC situation** – `LCs SITUATION - <date>.xlsx` (sheet `LC avec echéance`).
   * **Supplier invoices** – `Supplier Invoices - <month> W<n>.xlsx` (sheet `4411 - Parked - Downpayment`).
   * **UR Tracking Consolidation** – unrecognized revenue (`Summary UR`, `Sheet1` previous BU update, `Summary PO by BU`).
2. Extract the inputs (dates are the only things to type):
   ```bash
   python3 extract_inputs.py \
     --ar "AR_PRE_Week_4.xlsx" --prev-inputs inputs/history/weekly_inputs_2026-10-21.xlsx \
     --ar-baseline "AR_Beg_of_the_month.xlsx" --ar-baseline-date 2026-10-07 \
     --cash "CASH_FLOW_ANALYSIS - BUs October W4.xlsx" \
     --lc "LCs SITUATION - October 26th.xlsx" --ap "Supplier Invoices - Oct W4.xlsx" \
     --ur "UR Tracking Consolidation.xlsx" \
     --as-of 2026-10-28 --ar-date 2026-10-27 --ar-prev-date 2026-10-20 \
     --collections-to 26/10 --payments-to 22/10 --lc-date 26/10 --ap-label "Oct W4" \
     --out inputs/weekly_inputs.xlsx
   ```
   After the run, copy `inputs/weekly_inputs.xlsx` to `inputs/history/weekly_inputs_<as-of>.xlsx` so it can serve as
   next week's previous-week source. (`--ar-dashboard` accepts the old processed AR dashboard as the previous week; it was
   used once for the transition on 21 Oct 2026.)
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
| `AR_BU_DETAIL` | per BU: ownership, unbilled statuses, unpaid groups and phases, aging buckets, advance statuses | BU zoom slides |
| `AR_BU_PROJECTS` | per BU: top 8 unbilled and unpaid projects | BU zoom slides |
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
  `J-6-a Unearned on LS` and `C-3-a Pending DP clearing`; advance excludes `J-6-*` and `C-3-a` lines; at risk = `C-1-a/b`, SS postponed = `J-3-c`,
  under JESA = other `J-*`, under OCP = `C-*` and PDR). The aggregation reproduces the processed AR dashboard (`AR_DASH - COMPARE`) exactly.
* **Unrecognized revenue** = `Summary UR` grand total (all columns already include the 6% cost-and-fee markup);
  supplier cost on the slide = revenue ÷ 1.06. WoW and month-start references come from last week's deck
  (`--ur-prev` / `--ur-ms` JSON files by BU, or `--prev-inputs`).

## Conventions

* All amounts in MMAD (millions of MAD); AR decreases are shown green, increases red; cash net balances the other way round.
* Business-unit names are normalised (`AMT`→AM, `FEC`→Fertilizers & Chemicals, `WEN`→Water & Environment, …) in `extract_inputs.py`.
* The deck uses a text "JESA" wordmark top-right; drop the logo image on the slide master if preferred.

Dependencies: `python3`, `openpyxl`, `python-pptx` (`pip install openpyxl python-pptx`).
