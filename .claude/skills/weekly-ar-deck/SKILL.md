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
| `AR*.zip`, `AR*.csv` or `AR*.xlsx` | raw AR line list (sheet `AR`, or that sheet as CSV, zipped). On Google Drive it is the zipped CSV: the connector refuses files above 10 MB |
| cash flow workbook | **from the repository, not the drop folder**: `run_week.sh` pulls the newest `cashflow-bu/outputs/*/*.xlsx` (else `cashflow-bu/workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx`) from branch `claude/laughing-wright-x02plq`, maintained by the `cashflow-bu-update` skill. A `CASH*.xlsx` in the drop folder overrides it. If the extractor stops with "no cached values", the master was saved by a script: ask for the Excel-saved version or run the cash kit's delivery step |
| `AP*.xlsx` or `Supplier*.xlsx` | supplier invoices, sheet `4411 - Parked - Downpayment` (named `Invoices` from 10-Nov, same first 17 columns) |
| `*LC*SITUATION*.xlsx` | bank LC situation, sheet `LC avec echéance` |
| `UR*Tracking*.xlsx` | unrecognized revenue consolidation |
| `*Suspension*.xlsx` (when refreshed) | suspension log: project status joined to the unbilled lines (kept in `inputs/suspension_log/`) |
| `OCP*Payments*.xlsx` (when refreshed) | OCP payments weekly workbook: collections per OCP payment report by month and BU |
| `*DSO*.xlsx` (monthly) | DSO calculation workbook: 12-month trailing revenue per project, DSO trend |
| `*BU*Meeting*.pdf` (optional) | reconciliation reference (total AR, unbilled, unpaid, invoices paid) |

Download them into one scratch folder (the DSO workbook too when a new one is there). Drive downloads arrive as base64 JSON saved under the tool-results folder: decode the `content` field to the file named in `title`. A file above 10 MB cannot be fetched through the connector: stop and ask for the zipped-CSV version (AR) or the trimmed workbook (DSO), do not try to work around it. Do not run Python from inside that folder. If one of the five is missing,
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

## 2b. New month (first extract that carries the previous month's revenue)

The month-start baseline is the first AR extract of the month **that includes the previous month's revenue** (for
November 2026: the 10-Nov extract, not 03-Nov). On that run: copy the extract to
`inputs/baseline/AR_Beg_of_the_month_<YYYY-MM>.xlsx`, update `inputs/month_start.txt`, and run. Because baseline and
current extract are the same file, the deck drops the "vs MS" pills, columns and wording by itself and states on the
sources slide that month-start comparisons resume the following week. The executive summary and key message 2 isolate
the week-on-week movement on lines posted in the previous month's period ("of which +762 on lines posted in the October
period"): do not call it "revenue booked" and do not tie it to the P&L (some projects post cumulative figures).

## 3. DSO

A weekly `DSO Table` workbook (columns PU, 12 months trailing, Unbilled/Unpaid, DSO Exc ADV, Month revenue) takes
precedence when dropped (`*DSO*Table*.xlsx`): its DSO figures are used as published (BU meeting basis) and the previous
deck's DSO is shown as "last week". Without it, the rule below applies.

DSO is computed, not typed: total AR of the BU (from this week's AR extract) ÷ 12-month trailing revenue × 365.
The revenue side comes from the DSO calculation workbook (`WIP_DSO_calc.xlsx`, sheets `12 Months Trailing` and
`Trend`). If a refreshed workbook is in the drop (pattern `*DSO*.xlsx`), `run_week.sh` copies it to
`inputs/dso_calc/`; otherwise the latest one there is reused, which is correct until a new month of revenue is closed.
If the workbook is older than two months, say so in the summary. The extractor prints a NOTE when a computed DSO
differs from the workbook's own table by 2 days or more (the table may carry a manual override, e.g. JESA
Technologies 95 vs computed 106 in October); report it, do not change the calculation.

## 4. Run

```bash
cd weekly-ar-cash-deck
./run_week.sh <scratch folder> <as-of> <AR date> <previous AR date>
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
3. Upload the deck and the snapshot (pptx + png) to the delivery location (Google Drive folder `AR Cash Deck/Output`
   unless told otherwise) and send the files in the conversation when the reviewer asks for them, with the short
   cover email text of section 7.
4. Summary (short): the six headline figures with WoW and vs-month-start deltas, which source files were used and
   their dates, DSO source, warnings, anything left for a human (missing file, wording to confirm).

## 7. Cover email: one-page snapshot and short text

`run_week.sh` also builds `output/JESA_AR_Cash_Snapshot_<Mon>_W<n>_<as-of>.pptx` and `.png` with `build_onepager.py`
(same inputs). The PNG goes in the body of the cover email, the deck is attached. Layout agreed with the reviewer, do not
rearrange it:

- KPI boxes, in this order: gross exposure, unrecognized, total AR, unbilled, unpaid, DSO, cash received in the month
  (with the PSR / Proc split, and the prior month with its split), invoices billed in the month vs the prior month
  (billed = invoices deposited, by reception date, as in the BU meeting dashboard).
- Left panel: unbilled roll-ups, signed / not signed with the client (external) / not signed with JESA (internal).
- Middle panel: unrecognized revenue, equipment vs services totals, then signed vs not signed for equipment, services
  and total (bars drawn as shapes so every figure is readable).
- Right panel: "Cash position (Procurement Back to Back)" as a waterfall: cash IN, paid, approved suppliers, LC due,
  balance after approved, not approved, balance after all. The intermediate positive balance is not shown.
- Bottom: the previous-month revenue note (only in the weeks it applies) and the priority line.

Email text: short, bullet points only, amounts in MDH, no explanations (everything is in the deck). Template:

```
Please find attached this week's AR, Cash & Payables presentation.

Key figures as of <AR date>:
- Total AR: … MDH (… vs last week)
- Unbilled: … MDH (…% of AR)
- Unpaid: … MDH
- Unrecognized revenue: … MDH
- DSO: … days
- Collections <month> to date: … MDH
- Procurement cash: … MDH on actuals, … MDH after approved and pending payments
- Supplier payables: … MDH, of which … MDH past due

Roll-ups:
- Signed by the client: … MDH (…% of unbilled)
- Not signed: … MDH (PSR … MDH / Procurement … MDH)
- External, pending with the client: … MDH, mainly <top BUs>
- Internal, pending with JESA: … MDH, mainly <top BUs>
- Main internal blocker: … MDH <largest J status>

Priority remains to secure the client's sign-off on completed work (roll-ups / attachments) and to clear the internal
backlog, to unlock cash for PSR and critical suppliers.
```

## 8. Reviewer preferences (apply every week without being asked)

- Do not send files unless asked ("send me the presentation"); when the reviewer says "do not send", only confirm the change.
- No notes that reveal tooling: no file paths, script or guideline document names, no "previous slide" references.
- Say "client", not "OCP", for responsibility and signature wording. Abbreviations "WoW" and "vs MS" in key messages.
- Title slide: variance pills in MDH; when the AR increase comes from the previous month's revenue being posted, the
  note sits **below** the six boxes (automatic).
- The "Signed RU status" workbook (`Signed_RU_status.xlsb`) is a tracker of physical copies only: never use it for the
  signed / not signed totals. It can be quoted as a follow-up list of physical copies outstanding.
- Signed roll-up rule stays on the status mapping plus PA comments (`reference/status2_signed_mapping.md`).
- LC figures stay on the last bank LC situation received (19-Oct) until a new one is dropped; the LC payments due on
  the cash slide come from the COO cash file.
- Check whether the UR consolidation was really refreshed (compare totals by BU with last week); say so when it was not.
- If the COO cash file is re-sent during the week, re-run with it and report what moved (approved, pending, balances).
- Collections: AR report basis (all clients); explain the gap to the COO file as non-OCP clients and subsidiaries paid
  outside the OCP payment report when asked.

## Rules

- The signed / not signed and client / JESA responsibility classification of unbilled lines follows
  `weekly-ar-cash-deck/reference/status2_signed_mapping.md` (finance guideline). Change it only when finance changes the guideline.

- Cash position slide = the COO cash file headline, never overridden. When its by-BU table carries a "paid in the
  ongoing month" column, that amount is added to the approved payments (as in the file's headline) and mentioned on the card.
- The collections slide is on the AR report basis (all clients). It carries a one-line bridge to the OCP payment report
  amount used on the cash slide (difference = non-OCP clients and subsidiaries paid outside the OCP report).
- Never change the aggregation rules in `ar_engine.py` to make a number match; report the difference instead.
- Changes requested by the reviewer are applied in `build_deck.py` / `extract_inputs.py`, committed and pushed,
  so the next weekly run carries them.
- Amounts are MMAD (millions of MAD) everywhere; AR decreases are green, cash increases are green.
