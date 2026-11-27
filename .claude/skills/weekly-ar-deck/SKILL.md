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

## 2c. After a container reset (raw files gone)

The raw extracts (AR, baseline AR, LC situation, suspension log, DSO table, cash file) are not kept in git, and the
session container is rebuilt from the repository when it has been idle. Nothing has to be re-requested for:

- the **month-start baseline**: `run_week.sh` falls back to the saved inputs of the baseline week
  (`--baseline-inputs inputs/history/weekly_inputs_<date>.xlsx`, the file whose META `ar_current_date` is the baseline date);
- the **previous week**: always the latest `inputs/history/` workbook dated before the run;
- the **LC figures**: `--lc-inputs` reuses the LC lines saved in the previous inputs until a new bank situation is dropped;
- the **project statuses**: `inputs/suspension_log/status_map.json` (written each time a suspension log is processed).

Still needed every week from the reviewer: AR extract, COO cash file, supplier invoices, UR consolidation, OCP payments,
DSO table; the BU meeting PDF for the reconciliation; a suspension log when it has been refreshed.

## 3. DSO

When no DSO table is dropped, the extractor computes it: total AR of the BU (this week's extract, advances included, as
in the table) / 12-month trailing revenue carried from the last table x 365; the group figure uses the sum. A BU whose
published DSO does not follow that formula (JESA Technologies, manual basis) carries its last published figure.
Checked against the Power BI dashboard of 25-Nov-2026: identical for the group and all ten BUs. The revenue side only changes at month close: ask for a new table (or the month's revenue) after each close,
and check the computed figures against the Power BI dashboard when the reviewer shares it.

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
3. Signature safeguard: no `WARN roll-up status wording not recognised` in the extractor output, and the unsigned PSR by
   BU does not jump or sit far above what the BU reports (see Rules).
4. Read `KEY_MESSAGES` in `inputs/weekly_inputs.xlsx`: six messages in deck order, no internal or tooling references,
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
  (billed = invoices deposited = every line with a reception date in the month, whatever its status, as in the BU
  meeting dashboard).
- Under every KPI box: a strip with the PSR / Proc breakdown of that figure (unrecognized is all procurement; DSO
  shows "excl. advances").
- Left panel: unbilled roll-ups, signed / not signed with the client (external) / not signed with JESA (internal),
  each with its change vs last week in colour.
- Middle panel: unrecognized revenue, equipment vs services totals, then signed vs not signed for equipment, services
  and total (bars drawn as shapes so every figure is readable).
- Right panel: "Cash position (Procurement Back to Back)" as a waterfall with these labels: Cash IN, Paid, Ongoing
  payments (approved suppliers), LC due end of <month>, Net cash position, Payments pending approval, Final position.
  The intermediate positive balance is not shown. Below it: supplier payables and past due, each with its change vs last week.
- Bottom: the previous-month revenue note (only in the weeks it applies) and the priority line.

Email as actually sent (subject, wording and bullets are the reviewer's own; only the figures change):

```
Subject: JESA Cash Dashboard - <Month YYYY> (Week <n>)

Dear all,

Please find attached this week's AR, Cash & Payables presentation.

Key figures:

* Total AR: … MDH (… MDH vs last week)
* Unbilled: … MDH (…% of AR)
* Unpaid: … MDH
* Unrecognized revenue: … MDH
* DSO: … days
* Collections <month> to date: … MDH
* Cash Position: … MDH after approved and pending payments
* Supplier payables: … MDH, of which … MDH past due

Priority remains to secure the client's sign-off on completed work (roll-ups / attachments) and to clear the internal backlog, to unlock cash for PSR and critical suppliers.
```

The snapshot PNG goes in the body under the text; the deck is attached as `JESA GROUP Cash Position - <Mon YY>.pptx`
(`run_week.sh` writes that copy). No roll-up section in the email unless asked.

## 7b. Deck format as sent (default of `build_deck.py`)

The reviewer used to strip the deck by hand before sending; the builder now produces that version directly
(checked slide by slide against the deck sent on 18-Nov-2026). Do not put these back unless asked (`--full-notes` restores them):

- Key message on cash: headline only, no detail sentence.
- No explanatory footnotes on the unrecognized, collections (including the AR-report / OCP-report bridge) and
  payables slides; no "How to read" legend on the JESA status slide.
- Section dividers ("Business Unit Breakdown", "Backup") carry the number and title only.
- "Unbilled under JESA – Status by Business Unit" sits in the backup, first slide after the divider.
- No "paid and commitments" payments-detail slide and no "Sources & definitions" slide.

## 7c. Slides added on request (26-Nov-2026): keep them every week

- **AR Executive Summary (slide 4).** A band under each KPI box gives the PSR / Proc split of the box (gross exposure =
  AR PSR vs AR procurement + unrecognized). The payments card is a small grid: current month, prior month and since July,
  each with Total, PSR, Proc, OCP and Non-OCP (AR report, by payment date).
- **PSR – Cash In vs Revenue Posted by Business Unit** (right after the cash position slide). Revenue = rows
  "Revenue — PSR" of the cash flow workbook, sheet `REPORT - Cashflow by BU` (closed months since July). Cash in = PSR paid
  lines of the AR report by payment date, **all clients** (the OCP payment report alone understates Global Mining and
  JESA Technologies, mostly non-OCP). Coverage = cash in ÷ revenue over the same months. The month not yet closed has
  no posted revenue: the reviewer gives the **RFQ1** PSR revenue by BU; save it as
  `inputs/rfq1/psr_rfq1_<YYYY-MM>.json` (`{"month": "2026-11", "psr": {"AM": 60.0, ...}}`, MMAD) and the month joins
  the ratio, flagged "RFQ1". Without it the month's cash in is shown apart ("Nov cash in to <date>") and stays outside
  the ratio: ask for the RFQ1 figures, never estimate them.
- **Accounts Payable.** Cards: total payables incl. LC, open supplier invoices, past due (% of invoices), invoices under
  factoring (column "Does the supplier have a factoring?" of the supplier invoice file), LC opened. Table by BU with
  past due, % past due and factoring. Aging of the past due by days late (1–30, 31–60, 61–90, 91–180, > 180), factoring
  apart. Paid vs open invoices by supplier tier: paid in the closed months since July and ongoing payments of the
  current month from the tier rows of the JESA total block of `REPORT - Cashflow by BU`, open invoices and past due
  from the supplier invoice file ("Etranger" in the cash workbook = "International" not tiered in the invoice file).
- **Penalties on past due.** Shown only when the reviewer's penalties file has been received: save its figures as
  `inputs/penalties/penalties_<date>.json` (`{"total": MMAD, "label": "...", "by_bu": {"<BU>": MMAD}}`); a card and a
  "Penalties" column by BU then appear. Never compute or estimate penalties.

## 7d. "JESA Cash Risks" dashboard (stand-alone: executive slide + detail slide)

**Format as agreed with the reviewer (26/27-Nov-2026).** Slide 1:

- Headline: first priority payments until January, and penalties by end December and by end March if nothing is paid.
- First-priority box with the period ("until January <year>") and the waterfall, **largest category first**, each
  category bracketed with its total: supplier payments (AP) = past due factoring invoices + suppliers refusing (the AP
  team's "refusing" and "likely to refuse" merged, factoring left out) + tier 3 / old international invoices; LC due in
  the month in progress and the next two (November, December, **January**); customs of two months; penalties end
  December.
- Global exposure box "incl. Q1 <year> horizon": three invoice lines bracketed into "= Total AP in our books", LC
  outstanding, customs to March, **penalties to March** (accumulated Q1 estimate, not June).
- Row 1 panels: past due payables by age; factoring invoices past due by age (the part not yet due is a note only,
  not a bar); LC payments due.
- Row 2 panels: penalties by quarter (unpaid in blue, unbilled in gold); **"Critical payments"** = the waterfall as a
  table, largest category first, with the lines under supplier payments (AP): past due factoring invoices, "Suppliers
  refusing to pay penalties" (reviewer's wording for refusing + likely to refuse, merged), tier 3 and old
  international. LC is one category too, with one line per month underneath (November, December, January). The panel
  holds two tables side by side: supplier payments (AP) on the left; LC, customs and penalties on the right.
- No "decision needed" band (tried and removed at the reviewer's request).

Slide 2 "detail" = factoring by BU, penalties by payment date, customs paid and projected, unpaid invoices by supplier
position on invoice dates. No ledger account numbers (4411 / 4417) and "Q4", not "T4", on the slides. The notes below
describe the content of the blocks.

`build_deck.py --only risk --out output/JESA_Cash_Risks_<week>.pptx` (built by `run_week.sh`; `--with-risk` puts it in the
deck after the LC slide, only when asked). Content:

- Top left: **first priority payments**, with the waterfall: factoring past due, LC due in the month in progress, LC
  due next month (LC payment schedule, same figures as the LC panel; no sub-label under the LC bars), customs of the
  same two months (month in progress + next month at the average), critical invoices (total of the critical invoices
  section, factoring excluded so nothing is counted twice),
  penalties to be paid at end December (T4, from the penalties schedule), total first priority.
- Top right: **global exposure** box, no waterfall = supplier invoices (past due excl. factoring, not yet due excl.
  factoring, under factoring) + LC outstanding + customs from the month in progress to next March (paid to date in the
  month + projection; earlier months are not counted) + **penalties of the three payment dates up to June** (December,
  March and June; reviewer's instruction of 26-Nov-2026, replacing the earlier "T4 only").
- Panels, row 1: past due payables by age; factoring by BU (factoring, past due, not yet due); LC payments due over
  the next six months (do not mention LCs without a payment date or needing an extension: band = due in the two
  imminent months + after).
- Panels, row 2: customs paid by month plus the projection to March at the average of the closed months since July;
  penalties to be paid in December (T4), March and June as a small table, project / non-project invoices; **critical
  invoices to be paid first**, suppliers under factoring left out (they are paid in full as factoring past due), tab
  `AP_CRITICAL_GROUPS` for the two grouped lines and `AP_CRITICAL_INVOICES` for the invoices (one line per supplier and
  invoice reference):
  1. "Tier 3 national > 60 days late", one grouped line with its total;
  2. "International > 180 days late", one grouped line: international suppliers do not generate late-payment penalties,
     so they are treated apart and only come in when the due date is very old (penalty column "n/a");
  3. the national invoices that generate the most penalties (top three), ranked by amount x days past due until the
     penalties file gives penalties by invoice (`"by_invoice": {"<supplier>|<reference>": MMAD}`).
  The section total feeds the "Critical invoices" bar of the first-priority waterfall (`"priority_invoices"` in the
  penalties JSON overrides it).
- The global exposure box carries the label "incl. Q1 <year> horizon".
- Customs of the month in progress: paid to date by default; when the reviewer gives a full-month figure, pass it with
  `--customs-month <MMAD>` (26-Nov-2026: "make November customs fees 40 MDH"). It then feeds the first-priority
  waterfall, the global exposure and the customs panel, which quotes the amount paid to date next to it. It is a
  one-week figure: ask again the following week.
- Colours (reviewer, 26-Nov-2026: "less colourful, more in the JESA logo theme", then "a little more colour variant
  so that it's not faded"): JESA logo blue as the base (totals, table headers, first-priority box), a brighter blue
  (factoring), a light blue (not yet due, projections), a teal-blue (LC), slate grey (customs) and gold as the single
  warm accent (critical invoices, penalties, box title). No red or orange; past-due figures are in blue.
- Global exposure box: the three supplier-invoice lines (past due excl. factoring, not yet due excl. factoring, under
  factoring) are bracketed and followed by a sub-total line "= Total AP in our books" (reviewer, 26-Nov-2026: show
  that grouping the three categories gives the total AP; no separate slide for it).
- Each panel ends with a total band that ties it to the waterfall, in the colours of the bars (e.g. "Total past due
  1,309 = 1,007 excl. factoring + 302 under factoring"; LC = next six months + later + undated; customs = paid in the
  month + projection).
- The penalties boxes, the penalty-invoices bar and the penalties line of the global exposure stay blank ("–") and the total reads "before penalties" until the
  reviewer's penalties file is received; then add
  `"schedule": [{"due": "2026-12", "project": MMAD, "non_project": MMAD}, ...]` to the penalties JSON.

## 7e. Late-payment penalties estimate (AP team, "DDP")

Received 26-Nov-2026 by email from the AP team ("Delai de payment – risk in 2027"), three workbooks:

- `DDP estimate.xlsx`, sheet `All(Unpaid+unbilled)`: **accumulated** penalties by the end of Q4, Q1 and Q2 if the
  balances stay unpaid, for unpaid invoices (4411) and for goods / services received not yet invoiced (4417).
  26-Nov-2026: 57.4 / 93.3 / 127.1 MMAD. The slide shows the amount **per payment date** (increments: 57.4 in December,
  35.9 in March, 33.8 in June).
- `4411 DDP.xlsx`, sheet `All data`: penalty by SAP document and the AP team comment per supplier: `DDP` = refuses to
  change invoice dates (100% subject to penalties), `A risque` = may refuse, `Ok pour changement` = can change its
  invoice dates. The AP team asks to pay the first two before year end.
- `4417 DDP JESA & JEBI.xlsx`: detail of the unbilled side (not read; the total comes from the estimate workbook).

Pass `--ddp-estimate` and `--ddp-4411` (`run_week.sh` picks the latest `inputs/penalties/DDP_estimate_*.xlsx` and
`DDP_4411_*.xlsx`; raw files are git-ignored). Effects:

- Cash risks slide: penalties table (unpaid / not yet invoiced x December, March, June), "Penalties end Dec" bar of the
  first-priority waterfall = T4 amount, "Penalties to Jun" line of the global exposure = accumulated Q2 amount.
- Critical invoices section: after the tier 3 and international groups, the unpaid invoices of the suppliers
  **refusing** and **at risk**, with their penalty by end of December, and an "of which" line naming the three at-risk
  suppliers with the highest penalties. Suppliers under factoring (join on the SAP document number with the supplier
  invoice file) and the tier 3 group are left out of these two lines so nothing is counted twice. Suppliers that can
  change their invoice dates are not critical. The section total feeds the "Critical invoices" bar.
- Accounts Payable slide of the deck: card "Penalties end Q4 (est.)" with the Q1 and Q2 accumulated amounts.
- The file also covers invoices not related to a project (overheads): they are in the totals; about 33 MMAD of the
  4411 detail has no match in the supplier invoice file.
- Ask for a refreshed estimate when the payables move materially; quote it as an estimate, never recompute penalties.

## 8. Reviewer preferences (apply every week without being asked)

- Do not send files unless asked ("send me the presentation"); when the reviewer says "do not send", only confirm the change.
- No notes that reveal tooling: no file paths, script or guideline document names, no "previous slide" references.
- Say "client", not "OCP", for responsibility and signature wording. Abbreviations "WoW" and "vs MS" in key messages.
- Title slide: variance pills in MDH; when the AR increase comes from the previous month's revenue being posted, the
  note sits **below** the six boxes (automatic).
- **Signature rule, mandatory check every week.** For the PO, at-risk and VAT statuses the signature is read from the
  roll-up status column and the PA comments: "signed" / "signé" **or "Approved by Client PM"** = signed. Read the
  extractor's line `signature read from the roll-up status ...` and any `WARN roll-up status wording not recognised`:
  an unrecognised wording must be classified before the deck goes out (see the lesson in
  `reference/status2_signed_mapping.md`: this mistake overstated F&C's unsigned PSR by 229 MMAD in November 2026).
- A highlighted note with a clap can be shown on the PSR signature slide (`--psr-note "..."`, `**bold**` allowed), for
  an achievement to celebrate or a BU figure that runs ahead of SAP. It is a one-week note: do not carry it over.
- The "Signed RU status" workbook (`Signed_RU_status.xlsb`) is a tracker of physical copies only: never use it for the
  signed / not signed totals. It can be quoted as a follow-up list of physical copies outstanding.
- Signed roll-up rule stays on the status mapping plus PA comments (`reference/status2_signed_mapping.md`).
- LC figures stay on the last LC situation received until a new one is dropped (latest: LC follow-up of 17-Nov-2026,
  pass `--lc-date 17/11`); the LC payments due on the cash slide come from the COO cash file. Two layouts are read:
  the bank situation (opened and to-be-opened LCs) and the LC follow-up (opened LCs only, one Shipment / Payment
  column pair per month; the Payment columns give the schedule).
- Check whether the UR consolidation was really refreshed (compare totals by BU with last week); say so when it was not.
- If the COO cash file is re-sent during the week, re-run with it and report what moved (approved, pending, balances).
- Collections: AR report basis (all clients); explain the gap to the COO file as non-OCP clients and subsidiaries paid
  outside the OCP payment report when asked.

## Rules

- The signed / not signed and client / JESA responsibility classification of unbilled lines follows
  `weekly-ar-cash-deck/reference/status2_signed_mapping.md` (finance guideline). Change it only when finance changes the guideline.

- Cash wording, everywhere (cash slide, its backup, key message, snapshot): "Ongoing payments" (approved suppliers),
  "LC due end of <month>", "Net cash position" (after those two), "Payments pending approval", "Final position".
- Cash position slide = the COO cash file headline, never overridden. When its by-BU table carries a "paid in the
  ongoing month" column, that amount is added to the approved payments (as in the file's headline) and mentioned on the card.
- The collections slide is on the AR report basis (all clients). It carries a one-line bridge to the OCP payment report
  amount used on the cash slide (difference = non-OCP clients and subsidiaries paid outside the OCP report).
- Never change the aggregation rules in `ar_engine.py` to make a number match; report the difference instead.
- Changes requested by the reviewer are applied in `build_deck.py` / `extract_inputs.py`, committed and pushed,
  so the next weekly run carries them.
- Amounts are MMAD (millions of MAD) everywhere; AR decreases are green, cash increases are green.
