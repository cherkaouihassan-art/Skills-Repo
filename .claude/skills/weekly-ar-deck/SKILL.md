---
name: weekly-ar-deck
description: Build and deliver the weekly JESA "AR, Cash & Payables Position" PowerPoint from the five weekly source files (raw AR extract, cash flow workbook, AP information, LC situation, UR tracking consolidation) plus the BU meeting PDF for DSO. Use when asked to run, update or refresh the weekly AR / cash deck, or when a scheduled weekly run fires.
---

# Weekly AR, Cash & Payables deck – agent playbook

Everything lives in `weekly-ar-cash-deck/` (read its `README.md` first). The pipeline is deterministic:
`run_week.sh` → `extract_inputs.py` → `inputs/weekly_inputs.xlsx` → `build_deck.py` → `output/*.pptx`.
Your job is to fetch the files, set the dates, supply DSO, run, check, deliver.

## 1. Collect the week's files

Look, in this order, until you find the five files for the week:

1. the drop folder on the connected file source (Google Drive folder `AR Cash Deck/Inbox` unless the run prompt names another);
2. `weekly-ar-cash-deck/incoming/` in the repository (git-ignored; files committed there by the finance team);
3. files uploaded in the conversation.

Required (matched by name pattern, case-insensitive):

| Pattern | Content |
|---|---|
| `AR_PRE*.xlsx` | raw AR line list, sheet `AR` |
| `CASH*flow*.xlsx` | cash flow workbook, sheet `EXEC Summary` |
| `AP*.xlsx` or `Supplier*.xlsx` | supplier invoices, sheet `4411 - Parked - Downpayment` |
| `*LC*SITUATION*.xlsx` | bank LC situation, sheet `LC avec echéance` |
| `UR*Tracking*.xlsx` | unrecognized revenue consolidation |
| `*BU*Meeting*.pdf` (optional) | DSO pages |

Download them into one scratch folder. Do not run Python from inside that folder. If one of the five is missing,
stop and report which one; do not build a deck with stale data in its place.

Sanity checks before running: the AR file must be newer than the one archived in the last `inputs/history/` workbook
(compare total AR; identical totals mean the same extract was dropped twice), the AP totals must differ from last
week's `AP_KPI` values, and the LC file date in its name must be the latest available.

## 2. Dates

- `as-of` = the run date (YYYY-MM-DD).
- `AR date` = the AR snapshot date: the date in the drop folder's `dates.txt` if present (`ar_date=YYYY-MM-DD`),
  otherwise the working day before the run date.
- `previous AR date` = the `ar_current_date` stored in the `META` tab of the latest `inputs/history/` workbook.
- Month-start baseline: the raw AR file named in `inputs/month_start.txt` (path on line 1, date on line 2). The
  file itself is not in git (invoice-level data): fetch it from the file source folder `AR Cash Deck/Baseline` into
  `weekly-ar-cash-deck/inputs/baseline/` before running. If this is the first run of a calendar month, the AR file of
  this run becomes the new baseline: copy it to `inputs/baseline/AR_Beg_of_the_month_<YYYY-MM>.xlsx`, upload it to the
  `Baseline` folder, and update `month_start.txt`.

## 3. DSO

The BU meeting PDF cannot be parsed automatically. If a new PDF is in the drop, open its "DSO KPIs" / "DSO by BU"
pages with the Read tool (pages parameter) and write `inputs/dso_<AR date>.json` with the same keys as the previous
file (`JESA Group` plus each BU, `dso_current` and `dso_prior_month`). If no new PDF, reuse the latest `inputs/dso_*.json`
and say so in the summary.

## 4. Run

```bash
cd weekly-ar-cash-deck
./run_week.sh <scratch folder> <as-of> <AR date> <previous AR date> inputs/dso_<date>.json
```

Read the console output: `current AR`, `previous AR` and `month-start AR` lines, plus any `WARN` lines. A `WARN` on
the tracking log, priority list or LC ongoing sheets is acceptable (optional sheets); any other warning or a
traceback must be fixed before delivering (the extractor is label-anchored: compare the failing label with the sheet).

## 5. Check

1. Render: `python3 ../.claude/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf --outdir <scratch> output/<deck>.pptx`
   then `pdftoppm -r 60 -jpeg` and look at slides 1, 3, 4, 7, 8, 9, 12 and two BU zoom pages for overflow,
   overlapping text, "n/a" or zero-only charts.
2. Reconcile: total AR, unbilled, unpaid, advance and "invoices paid in month" must equal the BU meeting PDF
   when one was provided (the deck keeps DP-clearing lines in unbilled, like the PDF; the processed AR dashboard
   leaves them out).
3. Read `KEY_MESSAGES` in `inputs/weekly_inputs.xlsx`: six messages in deck order, no internal or tooling references,
   the cash message gives the net position after approved payments and after adding pending payments. Edit the
   wording if a figure reads oddly, then rebuild with `build_deck.py`.

## 6. Deliver and archive

1. `run_week.sh` already copied the inputs to `inputs/history/weekly_inputs_<as-of>.xlsx`; keep it (it is next
   week's previous-week source).
2. Commit `inputs/`, `output/*.pptx` and any code fix on the working branch and push.
3. Upload the deck to the delivery location (Google Drive folder `AR Cash Deck/Output` unless told otherwise) and
   send the file in the conversation.
4. Summary (short): the six headline figures with WoW and vs-month-start deltas, which source files were used and
   their dates, DSO source, warnings, anything left for a human (missing file, wording to confirm).

## Rules

- Never change the aggregation rules in `ar_engine.py` to make a number match; report the difference instead.
- Changes requested by the reviewer are applied in `build_deck.py` / `extract_inputs.py`, committed and pushed,
  so the next weekly run carries them.
- Amounts are MMAD (millions of MAD) everywhere; AR decreases are green, cash increases are green.
