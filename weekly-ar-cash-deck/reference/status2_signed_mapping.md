# Status 2 → responsibility and client signature (reference, finance guideline of 4 Nov 2026 – v2: at-risk lines also follow the roll-up status)

Responsibility: `Client` = action with the client / OCP, `JESA` = action with JESA. Signature: whether the attachment or service
sheet has been signed by the client. `Depends` = read the roll-up status column of the AR file: it is **signed** when that
column contains the word "signed" (and not "not signed"), otherwise **not signed**. `N/A` lines are not part of unbilled AR
(unpaid statuses, unearned on lump sum, advances) and are left out of the signed view.

| Status 2 | Responsibility | Signed or not | Note |
|---|---|---|---|
| C-1-a At risk (new project) | Client | Depends | roll-up status column |
| C-1-b At risk (budget overrun) | Client | Depends | roll-up status column |
| C-1-c Pending client PO | Client | Depends | roll-up status column |
| C-1-d Pending PO correction | Client | Depends | roll-up status column |
| C-2-a Waiting OCP VAT exemption | Client | Depends | roll-up status column |
| C-3-a Pending DP payment / clearing | Client | Signed | |
| C-3-b Invoicing waiting DP clearing | Client | Signed | |
| C-4-a Pending client PM approval | Client | NOT signed | |
| C-5-a Signed by client PM, pending OCP SS approval | Client | Signed | |
| C-5-b Not signed by client PM, pending OCP SS approval | Client | NOT signed | |
| C-6-a Pending SS OCP rejection | Client | NOT signed | |
| C-7-a Current | Client | N/A | unpaid |
| C-7-b Past due | Client | N/A | unpaid |
| J-1-a Pending JESA DMR / attachment creation PC | JESA | NOT signed | |
| J-1-b Pending attachments JESA PA | JESA | NOT signed | |
| J-1-c Pending LS milestone | JESA | NOT signed | |
| J-1-d Under JESA finance review | JESA | NOT signed | |
| J-1-e Under JESA PC review | JESA | NOT signed | |
| J-1-f Pending LC by PC | JESA | NOT signed | |
| J-2-a Pending JESA PM approval | JESA | NOT signed | |
| J-3-a Signed by client, pending JESA SS submission | JESA | Signed | |
| J-3-c SSS postponed | JESA | Signed | |
| J-4-a Ongoing invoicing | JESA | Signed | |
| J-5-a Rejected from client | JESA | N/A | usually under unpaid |
| J-6-a Unearned on LS | JESA | N/A | excluded from AR figures |
| J-6-b RG | JESA | N/A | excluded from AR figures |
| J-6-c Advance | JESA | N/A | excluded from AR figures |
| J-6-d ADVPAY | JESA | N/A | excluded from AR figures |
| J-7-a Under investigation | JESA | NOT signed | |

Unrecognized revenue: signed = the "attachment signed by client" columns of the UR consolidation (equipment + subcontracts).
The ownership view of the deck (under OCP / under JESA / SS postponed / at risk) is kept separately; this table drives the
"Signed roll-ups" slide only.
