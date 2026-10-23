# Weekly AR, Cash & Payables deck — JESA

A repeatable pipeline that turns four weekly source files into the executive PowerPoint
("AR, Cash & Payables Position – Weekly Executive Review").

```
sources (4 files)  ──►  extract_inputs.py  ──►  inputs/weekly_inputs.xlsx  ──►  build_deck.py  ──►  output/*.pptx
                                                 (review / edit by hand)
```

## Weekly input files (what to share each week)

| # | File | Sheet(s) read | Notes |
|---|---|---|---|
| 1 | `AR_PRE_Week_<n>.xlsx` | `AR` | raw AR line list of the week; drives AR, statuses, aging, non-OCP, JESA responsibility, collections, deposited |
| 2 | `CASH_flow_W<n>.xlsx` | `EXEC Summary` (+ `REPORT - Cashflow by BU`) | cash position, cash by BU, PSR collections. Optional sheets `Approved vs Tracking log …`, `Priority list …`, `LC Ongoing …` add the tracking-log and priority tables to the October payments backup slide; without `LC Ongoing` the month's LC maturities are taken from the LC situation file |
| 3 | `AP_information.xlsx` | `4411 - Parked - Downpayment` | supplier payables (downpayments, parked, 4411) – must be the extract of the current week |
| 4 | `<date> - LC_SITUATION.xlsx` | `LC avec echéance` | bank LC situation – latest version (maturity columns are read from the header) |
| 5 | `UR_Tracking_Consolidation.xlsx` | `Summary UR`, `Summary PO by BU`, `Sheet1` | unrecognized revenue (columns already include the 6% markup) |
| 6 | `WIP_DSO_calc.xlsx` | `12 Months Trailing`, `Trend`, `DSO Table` | DSO = total AR of the BU ÷ 12-month trailing revenue × 365. The AR side is recomputed every week from the AR extract; the revenue side (per project, trailing 12 months) only changes when a month closes, so a refreshed workbook is needed once a month (the last one is kept in `inputs/dso_calc/`). `Trend` gives the monthly DSO history for the trend slide. The BU meeting PDF is no longer needed for DSO |
| 7 | beginning-of-month AR extract | `AR` | once per month: the first AR extract of the month, path and date in `inputs/month_start.txt` |

Previous-week comparatives (AR, unrecognized revenue, AP, LC, collections) come from `inputs/history/weekly_inputs_<date>.xlsx`
saved by the previous run – nothing else to provide.

## Weekly run (about 5 minutes)

1. Put files 1–5 (and the DSO workbook when refreshed) in one folder (names can vary, they are matched by pattern:
   `AR_PRE*`, `CASH*flow*`, `AP*` or `Supplier*`, `*LC*SITUATION*`, `UR*Tracking*`, `*DSO*`).
2. Run everything in one command (extract → build → archive the inputs for next week):
   ```bash
   ./run_week.sh <folder> 2026-10-28 2026-10-27 2026-10-20
   #              sources   as-of      AR date    previous AR date
   ```
   The deck lands in `output/JESA_AR_Cash_Position_<as-of>.pptx` and `inputs/weekly_inputs.xlsx` is copied to
   `inputs/history/`.
3. Review `inputs/weekly_inputs.xlsx` if needed (`KEY_MESSAGES` wording, any manual correction), then rebuild with
   `python3 build_deck.py --inputs inputs/weekly_inputs.xlsx --out output/<name>.pptx`.
4. New month: save the first AR extract of the month and point `inputs/month_start.txt` to it (path, then date).

Manual equivalent of step 2 (all parameters explicit):
   ```bash
   python3 extract_inputs.py \
     --ar "AR_PRE_Week_4.xlsx" --prev-inputs inputs/history/weekly_inputs_2026-10-21.xlsx \
     --ar-baseline "AR_Beg_of_the_month.xlsx" --ar-baseline-date 2026-10-07 \
     --cash "CASH_flow_W4.xlsx" --lc "2026_10_26 - LC_SITUATION.xlsx" --ap "AP_information.xlsx" \
     --ur "UR_Tracking_Consolidation.xlsx" --dso inputs/dso_2026-10-20.json \
     --as-of 2026-10-28 --ar-date 2026-10-27 --ar-prev-date 2026-10-20 \
     --out inputs/weekly_inputs.xlsx
   python3 build_deck.py --inputs inputs/weekly_inputs.xlsx --out output/JESA_AR_Cash_Position_2026-10-28.pptx
   ```
   (`--ar-dashboard` accepts the old processed AR dashboard as the previous week; it was used once for the transition on
   21 Oct 2026. `--ur-prev` / `--ur-ms` JSON files override the unrecognized references when no history workbook exists.)

## Deck structure (41 slides)

1 Title · 2 Agenda · 3 Key messages · 4 AR executive summary · 5 Unbilled under JESA responsibility · 6 Non-OCP clients ·
7 Unrecognized revenue & suppliers · 8 Payments received · 9 Cash position · 10 Cash by business unit · 11 Accounts payable ·
12 Letters of credit · 13 "Business Unit Breakdown" divider · 14–31 two pages per business unit · 32 "Backup" divider ·
33 AR movement · 34 AR position by BU · 35 DSO trend · 36 Unbilled & unpaid status view · 37 PSR (labor) collections ·
38 October payments · 39 LC detail · 40 AP detail · 41 Sources & definitions.

Slides are built in code order and re-sorted by section at the end (`Deck.reorder` in `build_deck.py`); to move a slide
between the main body and the backup, change its `D.section = ...` tag.

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
| `COLLECTIONS_BU`, `COLLECTIONS_MONTHLY` | payments received by BU (month to date, prior month, YTD, non-OCP split, deposited) and monthly series | 4, 6, 8, 11, BU zooms |
| `CASH_MONTHLY` | group cash IN / OUT / net by month incl. October commitments, PSR block | 9 |
| `CASH_BU`, `CASH_BU_OCT` | cumulative and current-month cash position by BU | 13 |
| `CASH_PSR_BU` | PSR (labor) collections by BU and month, with the provision lines in `CASH_MONTHLY` | 14 (PSR slide) |
| `ONGOING` | paid / approved (LC, suppliers) / pending, tracking-log status, priority list | 9, 11 |
| `AP_BU` | payables & LC by BU, previous vs current week (AR dashboard) | 12 |
| `AP_BUCKETS`, `AP_BU_ORIGIN`, `AP_TIERS`, `AP_TOP_SUPPLIERS` | supplier-invoice file analysis | 12, 15 |
| `LC_LINES` | one line per LC with maturity schedule | 13, 14 |

## Comparatives

* **Week on week** = AR dashboard *Snapshot* tab (previous week) vs current.
* **vs month start** = beginning-of-month AR line list aggregated with the dashboard rules (unbilled excludes
  `J-6-a Unearned on LS`, DP-clearing lines `C-3-a` stay in unbilled; advance excludes `J-6-*` and `C-3-a` lines; at risk = `C-1-a/b`, SS postponed = `J-3-c`,
  under JESA = other `J-*`, under OCP = `C-*` and PDR). The aggregation reproduces the BU meeting report exactly; the processed AR dashboard (`AR_DASH - COMPARE`) differs only by the `C-3-a` lines it leaves out of unbilled (1.6 MMAD in week 3).
* **Unrecognized revenue** = `Summary UR` grand total (all columns already include the 6% cost-and-fee markup);
  supplier cost on the slide = revenue ÷ 1.06. WoW and month-start references come from last week's deck
  (`--ur-prev` / `--ur-ms` JSON files by BU, or `--prev-inputs`).

## Conventions

* All amounts in MMAD (millions of MAD); AR decreases are shown green, increases red; cash net balances the other way round.
* Business-unit names are normalised (`AMT`→AM, `FEC`→Fertilizers & Chemicals, `WEN`→Water & Environment, …) in `extract_inputs.py`.
* The deck uses a text "JESA" wordmark top-right; drop the logo image on the slide master if preferred.

Dependencies: `python3`, `openpyxl`, `python-pptx` (`pip install openpyxl python-pptx`).
