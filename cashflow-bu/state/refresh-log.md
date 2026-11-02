# Refresh log — CASH FLOW ANALYSIS - BUs

Newest first. One entry per refresh: inputs received, loads, headline movement, open questions.

## 2026-11-02 — October W4 refresh (first live refresh from Drive)
- **Inputs (Drive):** OCP Payments Weekly.xlsx (extract 26/10) · PAYMENTS OCTOBER W4 _ 29.10.26.xlsx · Payment tracking log.xlsx (02/11). LC and customs unchanged.
- **Loaded:** +37 OCP invoices (76.0 M) · +81 SAP payment lines (33.6 M included; 284 duplicates skipped; 24 ongoing lines closed, 6.9 M) · tracking-log check on 319 open lines (208 in log, 69 moved pending → approved, 83 PAID-in-log kept Open) · cover dates updated. 4 Data Sources rows added (102–105).
- **Headline (MAD):** collections 1,986,364,322 (67,114,841) · Oct actual -304,871,031 (-28,865,684) · approved -294,208,880 (-81,060,690) · pending -95,690,967 (87,915,665) · net all ongoing -257,404,853 (45,104,131).
- **Rulings:** 18 items (short/over-payments vs ongoing lines, 3 mixed-BU payments, 2 City of Lakeland payments without BU, 2 suppliers without tier) + 447 tracking-log lines not on the position (86.3 M). See `outputs/2026-11-02/movement-summary.md`.
- **Delivered:** `outputs/2026-11-02/CASH FLOW ANALYSIS - BUs October W4.xlsx`; master replaced.


## 2026-11-02 — baseline (October W3 as received)
- **Input:** user upload `CASH_FLOW_ANALYSIS_-_BUs_October_W3.xlsx` (11.0 MB, 59 tabs, 100 Data Sources rows, last load 23/10/2026).
- **Loaded:** nothing — baseline registered as master. Snapshot and shadow calculation run; 18/18 integrity checks OK; shadow headline ties to the cached headline to the MAD (1,268 report cells compared).
- **Headline (MAD):** collections 1,919,249,481 · dividends 185,098,380 · cash out Jul–Sep −1,734,096,677 · Oct actual −276,005,347 · approved −213,148,190 · net after approved −118,902,352 · pending −183,606,632 · net all ongoing −302,508,984.
- **As-of per stream:** OCP collections to 19/10 (File 26) · SAP payments to 15/10 (file 22/10) · tracking log W3 20/10 · ongoing lists to 23/10 (WEN Batch 5, AM list 22/10) · LC paid to 16/10 (file 21/10), LC situation 19/10 v003 · customs to 02/10 · revenue Jul–Sep.
- **Open questions:** see `refresh-state.json → open_questions`.
- **Next:** first refresh on receipt of OCP File 27+, SAP payments after 15/10, tracking log W4, LC situation, customs after 02/10.
