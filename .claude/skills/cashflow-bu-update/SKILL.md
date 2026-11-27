---
name: cashflow-bu-update
description: Refresh, reconcile or report on the daily "CASH FLOW ANALYSIS - BUs" workbook (JESA Group cash in / cash out by business unit). Use whenever the user mentions the cash flow analysis by BU, the cash flow report, OCP payment report collections, approved or pending supplier payments, the payment tracking log, LC payments, customs duties or the ongoing position, or asks for "an updated report". Covers intake of new inputs (files, emails, screenshots), loading them into the workbook, recalculation, integrity checks, the movement summary and the Data Sources audit trail.
---

# Cash flow analysis by BU — refresh skill

The master workbook, the input structure, state and templates live in `cashflow-bu/` at the repo root:

```
cashflow-bu/
  INPUT_REQUIREMENTS.md        the seven input streams, columns, keys, what we confirm back
  workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx   latest master (always refresh from this one)
  templates/CF_INPUT_PACK.xlsx intake pack: one tab per stream, for data that arrives without a file
  inbox/<yyyy-mm-dd>/          raw files received for a refresh (never edited)
  state/refresh-state.json     as-of date, source file and row count per stream, last headline
  state/refresh-log.md         one entry per refresh: inputs, headline movement, open questions
  outputs/<yyyy-mm-dd>/        delivered workbook + movement summary
```

Read `reference/workbook-map.md` (tabs, columns, formula criteria) and `reference/standing-rules.md` (business rules the user has set) before touching the file. The workbook also carries its own runbook on the hidden tab `HOW TO UPDATE`.

## Workflow for "update the report"

1. **Start from state.** Run `python3 scripts/snapshot.py cashflow-bu/workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx` to print the as-of dates, the headline and the integrity checks of the current master. Read `state/refresh-state.json` and the last entry of `state/refresh-log.md`. Tell the user what is loaded to which date.
2. **Collect inputs.** First look on Google Drive: `state/refresh-state.json → drive_sources` lists the folders and the file per stream (OCP weekly payments workbook, SAP payments extract, tracking log, LC situation, Inbox). Search each folder (`search_files` with `parentId = '<id>'`), take the newest file, `download_file_content` it (the result is persisted to a tool-results file for anything over a few hundred KB) and turn it into a file with `scripts/decode_drive_result.py <result.txt> cashflow-bu/inbox/<today>/`. Then classify everything received (Drive files, uploads, emails) into the seven streams of `INPUT_REQUIREMENTS.md`. For emails, screenshots and PDFs, transcribe the lines into a copy of `templates/CF_INPUT_PACK.xlsx` and keep the source reference in the `Source` column. Anything you cannot classify, ask — do not guess a BU or a tier. Note which Drive files are new since `drive_sources.seen_*` and which streams have nothing new.
3. **Normalise.** For each native file, map its columns to the intake tab layout (OCP `Payments Key`: Class `Procurement` → `Non-Labor`, `Proservices` → `Labor`; SAP W-files: `Type` is lower-case, BU codes `Water`/`F&C`/`Overhead` → canonical names) (header names in the pack match the source tabs). Canonical BU names and tiers are in `workbook-map.md` §3. Amounts: cash out negative, cash in positive, MAD. Keep the original reference exactly; the normalised key is computed by the loader.
4. **Load.** Native Drive files go through `scripts/build_pack_from_sources.py` first (OCP `Payments Key`, SAP payments with invoice-line BU mapping, tracking-log check tab `3b_LogCheck`); user decisions on existing lines go in a pack tab `8_Rulings` (actions `close`, `balance`, `pending`, `remove`, `include`, `bu`, `tier`, `master`, `master_remove`). Then: `python3 scripts/load_pack.py <master.xlsx> <pack.xlsx> --out <new master> --as-of <yyyy-mm-dd> --requested-by "<user>"`. It appends new rows to the right source tabs (dedupe on the stream key), extends the `OctOngoing` table, closes ongoing lines matched by SAP payments, writes one `Data Sources` row per stream, updates the cover date and sets Excel to recalculate on open. It prints what it added, what it skipped as duplicates and what needs a ruling. Use `--dry-run` first.
5. **Verify.** `python3 scripts/verify.py <new master> --baseline cashflow-bu/state/refresh-state.json`. The shadow calculation recomputes the whole report from the source tabs in Python (same SUMIFS logic, formula helper columns re-derived) and prints the new headline next to the baseline. Every headline move must be explained by a stream you loaded, using the loader's per-stream amounts. Add `--lo-scan` to also recalculate a *copy* in LibreOffice and scan the report tabs for `#REF!`/`#VALUE!` (several minutes); LibreOffice numbers are never used and that copy is never delivered, because LibreOffice 24.2 cannot evaluate the XLOOKUP/LET formulas. The delivered file recalculates in Excel on open (`fullCalcOnLoad`); when the user sends it back, `snapshot.py` reads Excel's own values and the 18 integrity checks.
6. **Deliver with values.** `python3 scripts/deliver.py <new master> <outputs/<today>/CASH FLOW ANALYSIS - BUs <label>.xlsx>`: it rewrites the report-feeding XLOOKUP/LET formulas as INDEX/MATCH (Excel and LibreOffice both evaluate them), recalculates a copy in LibreOffice so every cell carries a value (the file then shows figures in any viewer, phone preview included), re-enables Excel's full recalculation on open for the dynamic-array view tabs, and prints the snapshot (headline, 18 integrity checks) read from the calculated values. All 18 checks must read OK and the headline must equal the shadow calculation; otherwise stop and investigate. Never deliver the plain openpyxl-saved master: it has no cached values and looks empty outside Excel.
7. **Close the loop.** Update `state/refresh-state.json` and append to `state/refresh-log.md` (inputs, headline previous → new per stream, open questions). Save the new master over `workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx`, write `outputs/<today>/movement-summary.md`, commit.
8. **Report back** in chat: as-of per stream, headline table (previous / new / delta per stream), the open-questions list, and the path of the delivered file. Ask for the next inputs only if something is missing.

## Month roll-forward (first refresh with November actuals)
Follow `HOW TO UPDATE` section D in the workbook: new SAP tab `November Actual Payments` (same layout as October incl. helper columns AF:AK), insert a report column before the first Ongoing column, copy the October formulas and repoint the tab name, set row 4 to 2026-11-01, extend the cumulative formulas and the EXEC sections, re-run the checks. Rename the ongoing table convention (`OctOngoing` → keep the table name, update the header text) only if the user asks.

## Rules that must hold on every delivery
- Zero formula errors, 18 integrity checks OK, headline reconciled to the previous version.
- Never type numbers into black/green cells; never delete a source tab; never rename a BU header.
- Never delete ongoing lines — close them with `Manual status override`.
- Lines PAID in the tracking log stay Open until SAP shows the payment.
- Every load has a `Data Sources` row.
- No model identifiers or session links inside the workbook.

## Lean workbook (from 03/11/2026)
The user asked for a simplified file: `scripts/build_lean_workbook.py <full.xlsx> <lean.xlsx> --as-of <date>` builds `COO Cash Position` with a formula-only Summary and three value tables — `CashIn`, `CashOut` (SAP layout A:AC + Month/Kind/Include?/Cash out/Tier/BU/OH?/Source) and `Ongoing` (15 live columns) — plus Tiering Master, OH Suppliers and Log. It reconciles itself to the full workbook's cached headline and prints the differences. If the user adopts the lean file as the working file, future refreshes append to its three tables (same dedupe keys) instead of the full workbook's source tabs; keep the full workbook for BU detail on request.
