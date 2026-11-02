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


---
# Rulings applied 02/11/2026 (second load on the W4 master) — FINAL W4

| Line | W3 | W4 (Drive load) | W4 final (rulings) | Rulings move |
|---|---|---|---|---|
| Collections — OCP report (procurement) | 1,919,249,481 | 1,986,364,322 | 1,986,364,322 | 0 |
| Dividends (PROC compensable) | 185,098,380 | 185,098,380 | 185,098,380 | 0 |
| Cash OUT Jul–Sep actuals | -1,734,096,677 | -1,734,096,677 | -1,734,096,677 | 0 |
| Cash OUT Oct actuals | -276,005,347 | -304,871,031 | -304,709,453 | 161,578 |
| Approved payments (Oct ongoing) | -213,148,190 | -294,208,880 | -277,889,913 | 16,318,966 |
| Net balance after approved | -118,902,352 | -161,713,886 | -145,233,341 | 16,480,545 |
| Pending payment approval | -183,606,632 | -95,690,967 | -369,072,776 | -273,381,808 |
| Net position after all ongoing | -302,508,984 | -257,404,853 | -514,306,117 | -256,901,264 |

## What changed with your rulings
1. **Short / over-payments** — 9 ongoing lines closed: 3 where the payment exceeded the outstanding (PIM PERFECT, PROTECOR ×2) and 6 short-paid lines closed with a balance row for the difference (XY CONCEPT 54,622 · CARREFOUR INDUSTRIEL 69,700 · CRIP 10,904 and 6,365 · C.TRAV 9,994 · RESELEC 8,239; standing rule 9). The ISODEL/VOITH case was a reference collision, not a short payment, and was left alone.
2. **Mixed-BU payments** — dominant BU kept (PROTECOR → AM, ARVAL 4.67 M → AM, ARVAL 7,878 → B&I).
3. **CITY OF LAKELAND** — both payments (161,578) reclassified corporate, out of cash out.
4. **AM list 26/10** (`Approved AM Invoices Oct19th.xlsx`, requests of 26/10): 207 new lines added as approved (−64.3 M, 32 credit/reversal lines kept positive), 39 existing lines confirmed approved, 32 already on the position. The 125 lines flagged "Approved" 12/10 and 25/09 were all already loaded.
5. **Marine Maroc (Water)** — the 28 open Batch 5 invoices (4.12 M of the 4.31 M; 0.20 M was paid 13/10) are approved per the email of O. Zizi.
6. **Water program** — every other Water line that was not previously approved is pending: 24 Batch 5 lines that the tracking log had moved to approved go back to pending (78.6 M), and the WEN parked-invoice list of 19/10 adds 68 lines (195.0 M) as pending: 24 Batch 5 lines not yet on the position 78.4 M, 56 lines not yet in a batch 152.3 M, 12 lines marked *Future* 42.7 M. Previously approved Water lines (Batch 3, 20/10 reinstatements) stay approved.

## Open position after rulings (open lines, MAD)
- AM approved 246 lines −75.0 M · Water approved 40 lines −20.8 M · F&C approved 27 lines −6.0 M · B&I approved 3 lines −0.06 M
- Water pending 92 lines −273.4 M · F&C pending 27 lines −34.5 M · SP2M pending 2 lines −3.7 M · Energy pending 1 line −0.09 M

## Remaining questions
- The 12 Water lines marked *Future* on the WEN list (42.7 M, ATNER, SYCHEM…) are counted as pending because you asked for everything not previously approved; say if they should be excluded.
- Suppliers without a tier (loaded as Unclassified): HASLER GROUP, CONVEYOR PRODUCTS AND SOL (AM); ARDAMAN & ASSOCIATES, HANGZHOU WATER TREATMENT, FLUYTEC, TELSTAR MAROC (Water); REPUBLIC SERVICES, DBS BUILDING SOLUTIONS (SAP payments).
- Tracking-log lines not on the position: the AM block is now largely covered by the 26/10 list; corporate lines are out of scope.
- Still missing: LC payments after 21/10, LC situation after 19/10 (export under 10 MB), customs after 02/10.


---
# FINAL October W4 — after all rulings of 02/11/2026 (supersedes the two sections above where they differ)

| Line | W3 | W4 Drive load | **W4 final** |
|---|---|---|---|
| Collections — OCP report (procurement) | 1,919,249,481 | 1,986,364,322 | 1,986,364,322 |
| Dividends (PROC compensable) | 185,098,380 | 185,098,380 | 185,098,380 |
| Cash OUT Jul–Sep actuals | -1,734,096,677 | -1,734,096,677 | -1,734,096,677 |
| Cash OUT Oct actuals | -276,005,347 | -304,871,031 | -304,709,453 |
| Approved payments (Oct ongoing, incl. LC maturities) | -213,148,190 | -294,208,880 | -277,889,913 |
| Net balance after approved | -118,902,352 | -161,713,886 | -145,233,341 |
| Pending payment approval | -183,606,632 | -95,690,967 | -174,061,544 |
| Net position after all ongoing | -302,508,984 | -257,404,853 | -319,294,886 |

Corrections versus the first rulings pass: the WEN parked-invoice list captures **Batch 5 only**, so the 68 non-Batch-5 lines (195.0 M, incl. the 12 *Future* lines) were removed again (status `Removed - per user 02/11`). Water pending is the 24 Batch 5 lines not previously approved (78.4 M).

## Open commitments by BU (open lines, MAD; approved excludes the 143.9 M LC maturities shown separately on the report)
| BU | Status | Lines | Amount |
|---|---|---|---|
| AM | approved | 252 | -78,724,102 |
| Advisory | approved | 1 | 0 |
| Building & Infrastructure | approved | 15 | -2,355,076 |
| Energy Operations | approved | 2 | -15,768 |
| Energy Operations | pending | 61 | -7,065,320 |
| Fertilizers & Chemicals | approved | 47 | -9,744,379 |
| Fertilizers & Chemicals | pending | 27 | -34,525,258 |
| Global Mining | approved | 5 | -7,251,402 |
| Global Mining | pending | 4 | -10,692,041 |
| SP2M Ad Hoc Program | approved | 5 | -1,820,583 |
| SP2M Ad Hoc Program | pending | 7 | -43,408,348 |
| Water & Environement | approved | 49 | -31,804,840 |
| Water & Environement | pending | 24 | -78,370,577 |

## Tiering of the unclassified suppliers (applied; Tiering Master rows added, base "Research")
| Supplier | Tier applied | Basis |
|---|---|---|
| HASLER GROUP (AM, 3 lines) | Intl 1 | Tiering Master has HASLER France = OCP tier 1 and HASLER MAGHREB = tier 2; HASLER GROUP was paid as Intl 1 in July; French group |
| CONVEYOR PRODUCTS AND SOL (AM) | Not on OCP Tier | Master already lists "CONVEYOR PRODUCTS AND SOLUTIONS — Not on OCP Tier" (JESA base); Australian maker of conveyor rollers, Fenner/Michelin group. CPSM on the master is a different, Moroccan company |
| ARDAMAN & ASSOCIATES, INC (Water) | Intl 2 | Master: "ARDAMAN & ASSOCIATES INC" = OCP tier 2; US geotechnical firm (Orlando HQ, 14 branches) — international |
| HANGZHOU WATER TREATMENT (Water, 3 lines) | Intl 1 | Master: full name "Hangzhou Water Treatment Technology Development Center Co., Ltd." = OCP tier 1; paid as Intl 1 on LC payments |
| FLUYTEC, S.A. (Water) | Intl 3 | Master: "Fluytec; SA" = OCP tier 3; Spanish; paid as Intl 3 on LC payments |
| TELSTAR MAROC (Water) | 3 (estimate) | Not on the master. Kerix directory: Mohammedia, subsidiary of the Telstar group, under 10 staff, turnover 10–50 MDH → small national supplier |
| REPUBLIC SERVICES, INC · DBS BUILDING SOLUTIONS, INC (SAP payments) | corporate, out of cash out | US office running costs in Florida (waste collection in Lakeland; commercial cleaning in Tampa) — same treatment as CITY OF LAKELAND. Already flagged corporate by the invoice-line BU; no tier needed |
Tier moves between tier rows only; the headline does not change.

## Still outstanding
- LC payments after 21/10 and an LC situation newer than 19/10 (export under 10 MB); customs after 02/10.
- November roll-forward before November clearings are loaded.


---
# Cash IN update — OCP file of 02/11/2026
+53 invoices, 178.8 M: procurement 122.7 M (F&C 53.3 M, AM 43.2 M, Water 19.5 M, Global Mining 4.0 M, Energy 2.7 M) and PSR 56.1 M (Water 44.1 M, AM 8.9 M, Energy 3.1 M). October collections per OCP are now 468.8 M on 239 invoices. The single November invoice (Energy PSR 3.07 M, paid 02/11) is on the source tab and will show once the November column exists.

| Line | before | after | move |
|---|---|---|---|
| Collections — OCP report | 1,986,364,322 | 2,109,066,829 | +122,702,508 |
| Net balance after approved | -145,233,341 | -22,530,833 | +122,702,508 |
| Net position after all ongoing | -319,294,886 | -196,592,378 | +122,702,508 |
