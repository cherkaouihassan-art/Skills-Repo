# Scripts

| Script | Purpose | Typical call |
|---|---|---|
| `snapshot.py` | State of a workbook: controls, headline, 18 integrity checks (cached), as-of per stream, last Data Sources rows | `snapshot.py master.xlsx --json snap.json` |
| `build_template.py` | Rebuild the intake pack template | `build_template.py cashflow-bu/templates/CF_INPUT_PACK.xlsx` |
| `load_pack.py` | Load an intake pack into the master (append + dedupe, close matched ongoing lines, Data Sources rows, cover date, fullCalcOnLoad) | `load_pack.py master.xlsx pack.xlsx --out new.xlsx --as-of 2026-10-29 --requested-by "H." --label "October W4" [--dry-run]` |
| `shadow_calc.py` | Recompute the report from the source tabs and compare with cached values or a baseline | `shadow_calc.py new.xlsx --baseline cashflow-bu/state/refresh-state.json` |
| `verify.py` | LibreOffice-recalculate a copy, then shadow_calc on it (full cell-by-cell cross-check) | `verify.py new.xlsx --baseline cashflow-bu/state/refresh-state.json` |
| `cfcommon.py` | Shared constants (BU names, codes, tiers, tab names) and normalisers | imported |

Run times on the W3 master (11 MB): snapshot ~30 s, shadow ~40 s, load ~40 s, verify 5–10 min (LibreOffice).
Native source files (OCP workbook, SAP extract, tracking log export) are normalised into the pack layout first; the loader only reads the pack.
