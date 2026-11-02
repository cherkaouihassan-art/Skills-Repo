# Refresh log — CASH FLOW ANALYSIS - BUs

Newest first. One entry per refresh: inputs received, loads, headline movement, open questions.

## 2026-11-02 (night) — October closed, November ongoing, file delivered with values
- **Presentation:** October column labelled as closed actuals; the ongoing columns relabelled November (open commitments carried into November). No formula changed.
- **Delivery:** report-feeding XLOOKUP/LET formulas rewritten as INDEX/MATCH; file recalculated in LibreOffice so every figure shows in any viewer; Excel full-recalc-on-open kept for the view tabs. 18/18 integrity checks OK on the calculated values; headline = shadow calculation.
- **Delivered:** `outputs/2026-11-02/CASH FLOW ANALYSIS - BUs October closed + November ongoing (02-11).xlsx`.

## 2026-11-02 (evening) — cash IN from OCP file of 02/11
- **Input:** `OCP Payments Weekly 2.xlsx` (extract JESA-11.02.2026, uploaded). October now 239 invoices / 468.8 M (was 187 / 293.1 M at extract 26/10).
- **Loaded:** +53 invoices (178.8 M): procurement 122.7 M (F&C 53.3 M, AM 43.2 M, Water 19.5 M, GM 4.0 M, ENE 2.7 M) and PSR 56.1 M (Water 44.1 M, AM 8.9 M, ENE 3.1 M — the ENE line is dated 02/11 and sits outside the October column until the roll-forward). No unmapped invoices.
- **Headline (MAD):** collections 2,109,066,829 (+122,702,508) · net after approved -22,530,833 · net all ongoing -196,592,378.
- **Delivered:** `outputs/2026-11-02/CASH FLOW ANALYSIS - BUs October W4 final (cash in 02-11).xlsx`; master replaced.

## 2026-11-02 (final) — Batch 5 only + tiering
- **Rulings:** WEN parked-invoice list captures Batch 5 only → the 68 non-Batch-5 Water lines removed (195.0 M). Unclassified suppliers tiered from master variants / web research (HASLER GROUP Intl 1, CONVEYOR PRODUCTS Not on OCP Tier, ARDAMAN Intl 2, HANGZHOU Intl 1, FLUYTEC Intl 3, TELSTAR MAROC 3 est.); REPUBLIC SERVICES and DBS BUILDING SOLUTIONS corporate.
- **Headline (MAD):** approved -277,889,913 · pending -174,061,544 · net after approved -145,233,341 · net all ongoing -319,294,886.
- **Delivered:** `outputs/2026-11-02/CASH FLOW ANALYSIS - BUs October W4 final.xlsx`; master replaced. Data Sources rows up to 110.

## 2026-11-02 (later) — rulings on the W4 refresh → W4 final
- **Rulings (chat):** close lines when payment ≥ outstanding; short payments closed with balance rows; mixed-BU payments on dominant BU; CITY OF LAKELAND corporate (out); AM list of 26/10 approved (207 lines −64.3 M, credit lines kept positive); Marine Maroc Batch 5 approved per email (4.1 M open); Water: everything not previously approved is pending — 24 log-promoted Batch 5 lines reverted (78.6 M) and WEN parked list 19/10 loaded as pending (68 lines −195.0 M).
- **Headline (MAD):** approved -277,889,913 · pending -369,072,776 · net after approved -145,233,341 · net all ongoing -514,306,117 (W4 before rulings: -257,404,853).
- **Delivered:** `outputs/2026-11-02/CASH FLOW ANALYSIS - BUs October W4 (rulings 02-11).xlsx`; master replaced. Data Sources rows 106–109.

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
