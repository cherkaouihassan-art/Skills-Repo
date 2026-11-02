# CASH FLOW ANALYSIS – BUs — refresh kit

Daily/weekly cash-flow analysis of JESA Group by business unit (cash IN from the OCP payment report, cash OUT from SAP payments, approved and pending supplier payments, LC payments, customs duties).

| What | Where |
|---|---|
| Required inputs (the seven streams, columns, keys, what you get back) | `INPUT_REQUIREMENTS.md` |
| Latest master workbook | `workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx` |
| Intake pack for inputs that arrive without a file (emails, screenshots, PDFs) | `templates/CF_INPUT_PACK.xlsx` |
| Where to drop raw files for a refresh | `inbox/<yyyy-mm-dd>/` |
| Current as-of per stream, headline, open questions | `state/refresh-state.json` |
| History of refreshes | `state/refresh-log.md` |
| Delivered versions | `outputs/<yyyy-mm-dd>/` |
| Procedure and scripts Claude follows | `.claude/skills/cashflow-bu-update/` |

## How to get an updated report
1. Put the new files in `inbox/<today>/` (or upload them in the chat, or forward the approval email) — any mix of: OCP payments workbook, SAP payment extract, tracking log export, approved/pending lists or emails, LC payments / LC situation, customs detail.
2. Say "update the cash flow report".
3. You get back: the refreshed workbook, a movement summary (previous → new headline by stream), and the open questions that need your ruling (BU for unmapped payments, tier for new suppliers, log-paid-not-in-SAP lines).

The workbook's own runbook is on its hidden tab `HOW TO UPDATE`; the rules you have set over time are in `.claude/skills/cashflow-bu-update/reference/standing-rules.md`.
