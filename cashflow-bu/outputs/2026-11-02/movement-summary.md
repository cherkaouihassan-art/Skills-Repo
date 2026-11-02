# Movement summary — CASH FLOW ANALYSIS - BUs, October W3 (15/10–19/10) → October W4 (02/11/2026 refresh)

Sources (Google Drive, 02/11/2026): `OCP Payments Weekly.xlsx` (extract 26/10), `PAYMENTS OCTOBER W4 _ 29.10.26.xlsx` (SAP clearings to 29/10), `Payment tracking log.xlsx` (Invoice Tracker, actions to 02/11). LC situation (19/10) and customs (to 02/10) unchanged — no newer file.
Delivered file: `CASH FLOW ANALYSIS - BUs October W4.xlsx` (recalculates in Excel on open; 18 integrity checks to be read there). Headline below is the Python shadow calculation, which tied to Excel to the dirham on W3.

## Headline (MAD)
| Line | W3 | W4 | Move |
|---|---|---|---|
| Collections — OCP report (procurement) | 1,919,249,481 | 1,986,364,322 | 67,114,841 |
| Dividends (PROC compensable) | 185,098,380 | 185,098,380 | 0 |
| Cash OUT Jul–Sep actuals | -1,734,096,677 | -1,734,096,677 | 0 |
| Cash OUT Oct actuals | -276,005,347 | -304,871,031 | -28,865,684 |
| Approved payments (Oct ongoing) | -213,148,190 | -294,208,880 | -81,060,690 |
| Net balance after approved | -118,902,352 | -161,713,886 | -42,811,534 |
| Pending payment approval | -183,606,632 | -95,690,967 | 87,915,665 |
| Net position after all ongoing | -302,508,984 | -257,404,853 | 45,104,131 |

## Why it moved
- **Collections +67.1 M** — 37 new OCP invoices paid 20–26/10 (76.0 M incl. 8.9 M PSR, which sits outside the procurement headline): AM procurement 46.8 M, Water procurement 20.3 M, AM PSR 8.8 M, JET PSR 0.1 M. September ties to the OCP summary exactly.
- **Oct actual cash out −28.9 M** — 81 new SAP payment lines (33.6 M included) cleared 16–29/10: Water 14.8 M, AM 11.7 M (of which ARVAL 4.7 M flagged OH and added back), Global Mining 6.8 M (CLARIANT 6.7 M). 284 lines already loaded were recognised as duplicates.
- **Approved −81.1 M** = 24 ongoing lines closed by those payments (+6.9 M: 21 AM lines, CA2E / SIDEN / MARINE MAROC for Water) and 69 lines that moved from pending to approved because they are now in the tracking log (−87.9 M, of which Water Batch 5 85.2 M, SP2M 1.8 M, B&I 0.8 M).
- **Pending +87.9 M** — the same 69 lines leaving pending.
- **Net after all ongoing +45.1 M** = +67.1 (collections) −28.9 (payments) +6.9 (commitments cleared).

## Rulings needed (18)
- 2_CashOut_SAP: invoice 'F180-2026' (XY CONCEPT) paid 273,110 on doc 2000012418 vs ongoing row 1353 amount 327,732 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice 'FA22556' (RESELEC) paid 74,150 on doc 2000012416 vs ongoing row 1092 amount 82,389 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice '037/2026' (CRIP) paid 76,328 on doc 2000012415 vs ongoing row 1085 amount 87,232 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice '026/2026' (CRIP) paid 44,553 on doc 2000012414 vs ongoing row 1108 amount 50,918 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice 'FE263324' (CARREFOUR INDUSTRIEL ET) paid 627,297 on doc 2000012412 vs ongoing row 160 amount 696,997 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice 'CT046/26' (C.TRAV) paid 109,936 on doc 2000012341 vs ongoing row 1012 amount 119,930 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: payment 154,169 PROTECOR doc 1500039282 — mixed BUs on invoice lines: AM 219,263, Water & Environement 30,834 — dominant AM used
- 2_CashOut_SAP: payment 4,668,675 ARVAL MAROC doc 1500039276 — mixed BUs on invoice lines: AM 3,658,079, Overhead (corporate) 1,861,512, Global Mining 43,105, Fertilizers & Chemicals 19,764 — dominant AM used
- 2_CashOut_SAP: invoice 'FA009/2026' (PIM PERFECT INSTALLATION) paid 186,226 on doc 1500039217 vs ongoing row 809 amount 170,445 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice '16/2026' (PROTECOR) paid 20,476 on doc 1500039216 vs ongoing row 1147 amount 17,063 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: invoice '15/2026' (PROTECOR) paid 57,549 on doc 1500039216 vs ongoing row 1109 amount 47,957 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: payment 65,020 PROTECOR doc 1500039216 — mixed BUs on invoice lines: AM 92,473, Water & Environement 13,004 — dominant AM used
- 2_CashOut_SAP: invoice 'FA2600157' (ISODEL SARL) paid 12,153 on doc 1500039215 vs ongoing row 1132 amount 23,040 — partial/short payment: confirm closing and balance row
- 2_CashOut_SAP: payment 7,878 ARVAL MAROC doc 1500002660 — mixed BUs on invoice lines: Building & Infrastructure 9,462, Water & Environement 5,244 — dominant Building & Infrastructure used
- 2_CashOut_SAP: supplier 'REPUBLIC SERVICES, INC' has no tier on Tiering Master — tier needed
- 2_CashOut_SAP: payment 61,098 CITY OF LAKELAND, FLORIDA doc 1500001534 (invoices 3637783 09/26) — no BU found: give a BU (counted in consolidated Unclassified until then)
- 2_CashOut_SAP: supplier 'DBS BUILDING SOLUTIONS, I' has no tier on Tiering Master — tier needed
- 2_CashOut_SAP: payment 100,480 CITY OF LAKELAND, FLORIDA doc 1500001524 (invoices 3637783  01/26, 3637783 08/26, 3637774 07/26) — no BU found: give a BU (counted in consolidated Unclassified until then)

## Tracking log lines not on the ongoing position
447 unpaid lines in the Invoice Tracker (86,273,466 MAD) have no line on the ongoing position. Not loaded (the log confirms approval, it does not create commitments without your ruling). By BU and status:
- AM | PAID,: 1 lines, 109,544 MAD
- AM | PENDING AP: 218 lines, 59,802,635 MAD
- AM | PENDING TRESO: 49 lines, 10,724,798 MAD
- Building & Infrastructure | PENDING AP: 2 lines, 79,080 MAD
- Corporate | : 8 lines, 0 MAD
- Corporate | PAID,: 16 lines, 88,825 MAD
- Corporate | PENDING AP: 126 lines, 4,268,230 MAD
- Corporate | PENDING TRESO: 7 lines, 1,177,169 MAD
- Fertilizers & Chemicals | PAID,: 4 lines, 8,078,057 MAD
- Fertilizers & Chemicals | PENDING AP: 10 lines, 1,300,551 MAD
- Fertilizers & Chemicals | PENDING TRESO: 4 lines, 610,329 MAD
- SP2M Ad Hoc Program | PENDING AP: 1 lines, 18,078 MAD
- Water & Environement | PAID,: 1 lines, 16,170 MAD

Corporate lines are outside the procurement scope anyway; the AM PENDING AP block (218 lines, 59.8 M) is the one to rule on: add to AM approved, or leave to the AM lists.

## Still outstanding inputs
- LC payments after 21/10 and a newer LC situation (the 19/10 file on Drive is 11 MB, above the connector's 10 MB download limit — a smaller export or the paid-LC file is needed).
- Customs detail after 02/10.
- Pending-approval lists received after 23/10 (none on Drive).
