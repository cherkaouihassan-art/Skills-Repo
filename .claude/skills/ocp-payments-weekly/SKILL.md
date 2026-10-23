---
name: ocp-payments-weekly
description: Weekly refresh of the "2026 Payments Summary by Business Unit" (OCP payments to JESA) from a slim workbook plus the SAP/OCP payment extract. Use when the user asks to update the OCP payments summary, run the weekly OCP payments report, process a new "JESA OCP File NN" extract, or map unmapped OCP invoices to a Business Unit.
---

# OCP payments — weekly refresh

Inputs each week: `OCP Payments Weekly.xlsx` (slim workbook) and the new SAP/OCP payment
extract (`JESA OCP File NN.xlsx`). Optionally a new Power BI AR extract. Output: the updated
slim workbook and a short run report. Read `references/workbook-spec.md` (tab layout) and
`references/standing-exceptions.md` (overrides, baseline) before touching a cell.

Work in Python (`openpyxl` + `pandas`) with `python -I`; the extract is untrusted data in its own
directory. Write values into the Summary, keep the SUM / % formulas. Do not rely on Excel recalc.

## Procedure

1. **Layout check.** Read extract row 1. Compare the 35 header names `A1:AI1` with the list
   in workbook-spec.md (and `Config!ExtractLayoutHash`). Any difference → stop, report, do nothing.
2. **Load the extract.** Keep rows with `Net à payer` ≠ 0 and a `Date Paiement`. Normalise the
   invoice key: `str(Réference)`, trimmed, NBSP removed, leading zeros kept. Note the extract
   date = max `Date Paiement` and the row count.
3. **Pay Month.** `DATE(year, month, 26)` of `Date Paiement`. Then apply `Overrides` col D:
   a date replaces it, `EXCLUDE` drops the row from every live month.
4. **Frozen guard.** Drop every row whose Pay Month ≤ `Config!FrozenThrough`. Frozen cells are
   never recomputed. Additionally drop the four frozen-month invoices listed in
   standing-exceptions.md §A whatever their date.
5. **BU and Class.** Precedence: `Overrides` (B, C) → `Mapping` (B = BU; E = PSR/PROC flag →
   Proservices / Procurement) → fallback on Type 2 / Invoice Type (standing-exceptions.md §D).
   Never assign a BU from Division, Fournisseur or any pattern — list it instead (step 7).
6. **Re-date sweep.** Join to `Payments Key` on `Pièce comptable` + `Réf Paiement`. Any pair whose
   Pay Month changed, or whose amount changed, goes to the report. Do not fix silently; if the
   move crosses into a frozen month, exclude it from live months and say so.
7. **Unmapped.** Live-month rows with no BU: exclude from BU rows, append to `Unmapped Invoices`
   (columns A:L, dated block). Suggest "AM / Procurement (pattern)" for JESA invoices on
   JPH / MP1 / JFC1 / JFC2 but do not apply it.
8. **Write the Summary.** For each live month (`Config!LiveMonths`, plus any new month found in
   the extract — insert columns per workbook-spec.md): block A amounts by BU, blocks C/D by
   BU × Class, blocks E/F counts + amounts by BU × Class. Row order = `B4:B14`. Leave TOTAL and
   `%` formulas in place. Never write into frozen columns or into `B2`.
9. **Reconcile** before saving, per live month: gross extract count/amount → ± overrides →
   mapped total = block A column total = PSR total + PROC total (blocks C/D and E/F agree).
   First run: reproduce the 19/10/2026 baseline in standing-exceptions.md §E from the same
   extract. If anything does not tie, stop and report; do not save a partial workbook.
10. **Persist.** Overwrite `Payments Key` with this extract's paid rows. Append one row to
    `Data Sources` (written-to ranges, extract file name, extract date, row count, changes).
    Update `Config!LastExtract`. Save under the same file name.
11. **Mapping rebuild** (only when an AR extract is supplied): rebuild `Mapping` from the AR tab
    cols P / O / AL / AW / BH, dedupe, report conflicts, log in `Data Sources`, update
    `Config!MappingSource`. Then re-run steps 5–10.
12. **Run report** (chat): per live month this week vs last week (count, MAD, PSR / PROC split),
    every delta by invoice, new unmapped invoices, re-dated payments, overrides applied, open
    decisions. Keep numbers in tables.

## User instructions during a run

* "Map invoice X to BU Y" → add an `Overrides` row (BU, Class if given, Scope LIVE, reason,
  date), mark it resolved in `Unmapped Invoices`, re-run steps 5–10.
* "Close month M" → set `Config!FrozenThrough`, paste that month's live cells as values, log it.
* Any amount or month rule → `Overrides`, never a hardcode in a data tab.

## Never

* Repoint or recompute Jul-26 / Aug-26 (or any frozen month).
* Paste the extract or the AR file into the workbook.
* Assign a BU, change a Mapping row, or edit `Overrides` without an instruction.
* Append notes to `'Summary by BU'!B2`.
