# Standing rules — captured from the workbook's own audit trail (Data Sources rows 2–101) and user rulings

Apply these on every refresh unless the user overrules them in writing. When the user overrules one, record the ruling on `Data Sources` and update this file.

## Signs, units, scope
1. Cash OUT is negative, cash IN positive, everything in MAD. Credit notes on the ongoing position are negative too (they reduce cash out) — per user instruction 09/09.
2. FX for EUR invoices on the ongoing position: 10.8 (ex-M1 default) or 10.66 (F&C September list); note the rate used in column AE/AF.
3. Only the 10 mapped BUs are reported. Unmapped / corporate / Agency Local / Filiale lines stay out of the BU blocks and are listed on `Oct Pmts - Not BU-mapped`.

## Cash IN (OCP payment report)
4. Source of truth is the OCP payment report (`JESA OCP File nn`); the report must tie exactly to the OCP "Summary by BU" per BU × month × PSR/PROC after every load.
5. Labor vs Non-Labor comes from the Power BI - AR flag; fallback is the AR mapping. Standing overrides: AM re-mapping exceptions of 06/10 (5 invoices) and invoice 72000109 (Water advance, manual payment date 07/10) stay until reversed.
6. Dividends (PROC compensable) are a standing manual layer, never refreshed from a payment file.
7. PSR collections of AM are shown but excluded from the procurement headline (ring-fenced for AM PSR costs). Water PSR reserve 100 MMAD is a provision, not cash.

## Cash OUT actuals (SAP)
8. Only `Paiement` lines with `Paiement positive = Oui` count; invoice lines, reversals (same-day clearing reversed) and excluded tiers are flagged `No - …` in the Include column, never deleted.
9. A SAP payment closes an ongoing line only when matched on reference (then supplier + amount). Short payments create a balance row (commitment − net cash) on OctOngoing.
10. Payments found in SAP that are not on the ongoing list are listed for the user to attribute a BU before they enter a BU block.
11. Customs = Treasury "RECEVEUR DOUANE / TGR" detail by BU and date (tab `DD`) — not the SAP supplier 8138 lines (checked for double counting 14/09: none).

## Ongoing position (approved vs pending)
12. **Approved** = line is in the Payment tracking log (Invoice Tracker) OR on the user's approved list (Priority list, AM approved list W1, Energy Ops approved list, WEN batches…). **Pending approval** = everything else still open.
13. A line marked PAID in the tracking log stays **Open** until it appears in a SAP payment extract (standing rule, 20/10). Log-paid-but-not-in-SAP lines are chased with Treasury, not closed.
14. Lines are never deleted: close them with `Manual status override` (`Paid in Oct`, `Superseded - …`, `Deferred - …`, `Removed - per user dd/mm`) and keep the comment trail in column I.
15. A new BU list supersedes the previous list for that BU only when the user says so (Water Batch 3 → 25/09 position → Batch 4 → Batch 5); otherwise de-duplicate on reference and add only new lines.
16. Tier comes from `Tiering Master` (OCP database first); unknown suppliers go in as `Unclassified` and are listed for the user to tier.
17. Overhead suppliers (`OH Suppliers`) are flagged `OH` and excluded from the procurement headline for AM and Water, kept as memo for other BUs.
18. AM PISTELLI / PELZ is excluded from AM cash out (cash-in recovery), Intl 2.

## LC
19. LC actuals come from the bank statement LC lines (Treasury "PAIEMENTS PAR LC" file), not from the drawdown schedule (schedule superseded 29/09).
20. LC ongoing = October maturities per the latest bank LC situation ("LC avec echéance"); an LC already paid in the month (reference match with SAP or the LC payments file) is set to 0 in column Q. Each new situation gets a dated tab and the formulas are repointed.

## Process
21. Every load gets a `Data Sources` row (source file, object, parameters, run time, row count, changes after retrieval, effect on the headline).
22. After every load: full recalculation, all 18 integrity checks must read OK, and the headline move must be explained in one line per stream.
23. Deliver the workbook plus the one-page movement summary (previous headline → new headline by stream).

## Rulings of 02/11/2026
24. A SAP payment equal to or above the ongoing line clears the line. A payment below it closes the line and carries the difference as a balance row (rule 9).
25. When a payment's invoice lines span several BUs and the cash-out file gives no BU, use the dominant BU by amount.
26. US corporate charges (CITY OF LAKELAND etc.) are corporate, not part of cash out (`No - corporate OH`).
27. Water program: only lines previously approved by the user (Batch 3 approved request, 20/10 reinstatements, Priority list, Marine Maroc per O. Zizi email) are approved; every other Water line received (WEN parked-invoice list, Batch 5 and later) is pending approval even when it appears in the tracking log.
28. AM: the AM approved list (requests dated 26/10 in `Approved AM Invoices Oct19th.xlsx`) is approved. Positive lines in supplier lists are credit notes / reversals and keep their sign.
29. SAP payments match ongoing lines on the invoice references of the payment's invoice lines (same clearing document), at net cash (invoice TTC less withholding lines).
30. The WEN "4411 parked invoices / downpayment" list is used for its Batch 5 lines only; lines marked FALSE / 0 / Future / paid batches are not commitments on the position.
31. Tiering of a supplier absent from the master: look for the same company under another spelling (HASLER France / HASLER GROUP, Fluytec; SA / FLUYTEC, S.A., the long Hangzhou name), then web research for nationality and size; record the result on the Tiering Master with base "Research" and the reason in Detail. US office running costs (Lakeland / Tampa suppliers) are corporate.
