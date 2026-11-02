# Scripts used by the weekly run (python3 -I, openpyxl)

All scripts expect the slim workbook at `../out/OCP Payments Weekly.xlsx` relative to the
working directory and write a `weekly_state.pkl` next to themselves. Run from a scratch
directory; keep the extract and AR files in separate input directories.

1. `weekly_dryrun.py <workbook> <extract.xlsx>` — layout check, classification
   (Overrides > Mapping > fallback), re-date sweep against `Payments Key`, unmapped list,
   per-month pivots. Prints the reconciliation; saves `weekly_state.pkl`. Nothing is written.
2. `weekly_write.py` — env vars `EXTRACT_NAME`, `RUN_DATE` (YYYY-MM-DD), `RUN_NOTE`, `RUN_TAIL`.
   Inserts a column group for any new live month (blocks A/C/D before Total, E/F appended),
   writes live cells as values, rewrites `Payments Key`, appends `Unmapped Invoices`, logs to
   `Data Sources`, updates `Config`. Prints week-over-week new payments.
3. `mapping_rebuild.py <AR.xlsx> <workbook>` — only when an AR extract is supplied. Rebuilds
   `Mapping` (first occurrence, carry-over of missing invoices, blank flags filled), reports
   BU / flag changes on live-month invoices, logs to `Data Sources`. Re-run 1 and 2 afterwards.

After step 2 always recalc with LibreOffice (`soffice --headless --convert-to xlsx`) and check:
grand total = sum of month totals, each month = PSR + PROC, blocks E/F = blocks C/D.
