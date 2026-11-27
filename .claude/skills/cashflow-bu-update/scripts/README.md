# Scripts

| Script | Purpose | Typical call |
|---|---|---|
| `snapshot.py` | State of a workbook: controls, headline, 18 integrity checks (cached), as-of per stream, last Data Sources rows | `snapshot.py master.xlsx --json snap.json` |
| `build_template.py` | Rebuild the intake pack template | `build_template.py cashflow-bu/templates/CF_INPUT_PACK.xlsx` |
| `load_pack.py` | Load an intake pack into the master (append + dedupe, close matched ongoing lines, Data Sources rows, cover date, fullCalcOnLoad) | `load_pack.py master.xlsx pack.xlsx --out new.xlsx --as-of 2026-10-29 --requested-by "H." --label "October W4" [--dry-run]` |
| `shadow_calc.py` | Recompute the report from the source tabs and compare with cached values or a baseline | `shadow_calc.py new.xlsx --baseline cashflow-bu/state/refresh-state.json` |
| `verify.py` | shadow_calc against the baseline (+ optional `--lo-scan`: LibreOffice recalculates a copy only to scan for #REF!/#VALUE!) | `verify.py new.xlsx --baseline cashflow-bu/state/refresh-state.json [--lo-scan]` |
| `deliver.py` | Build the deliverable with calculated values: INDEX/MATCH rewrite → LibreOffice recalc → full-calc-on-load flag → snapshot | `deliver.py new.xlsx outputs/<date>/<name>.xlsx` (5–10 min) |
| `make_lo_compatible.py` | The XLOOKUP/LET → INDEX/MATCH rewrite used by deliver.py | called by deliver.py |
| `build_pack_from_sources.py` | Native Drive files → intake pack (OCP Payments Key, SAP payments with invoice-line BU, tracking-log check) | `build_pack_from_sources.py master.xlsx --template … --out pack.xlsx --ocp f --sap f --log f` |
| `decode_drive_result.py` | Persisted Drive download result → file | `decode_drive_result.py result.txt inbox/<date>/` |
| `cfcommon.py` | Shared constants (BU names, codes, tiers, tab names) and normalisers | imported |

Run times on the W3 master (11 MB): snapshot ~30 s, shadow ~40 s, load ~40 s, verify ~40 s (+5–10 min with --lo-scan).

`shadow_calc.py` was validated on the October W3 master: 1,268 report cells and the eight headline figures tie to Excel's cached values to the MAD. On a file saved by openpyxl (no cached values) it re-derives every formula helper column (ongoing status / key / OH / approved, September BU / tier / OH, LC situation BU / paid flag, customs BU / month, dividends) and compares the headline only.
Native source files (OCP workbook, SAP extract, tracking log export) are normalised into the pack layout first; the loader only reads the pack.
