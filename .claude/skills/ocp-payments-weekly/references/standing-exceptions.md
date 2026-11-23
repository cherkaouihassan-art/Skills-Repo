# Standing exceptions (seed rows for the `Overrides` tab)

Verified against `OCP_Payments_report.xlsx` of 19/10/2026 (`'JESA OCP File 26'`, hardcoded
cells in `AJ:AL`, the `B2` audit note and the `Data Sources` log). These rules win over the
extract and over `Mapping`. Nothing is added here without an explicit user instruction.

## A. Amount / month rules on invoices already in frozen months (Scope = FROZEN)

| Invoice | Amount (MAD) | Where counted | Rule for live months |
|---------|--------------|---------------|----------------------|
| 91000988 (OCP SA, acompte recovery) | 110,500,000 at AR value | Aug-26, frozen | Live extracts carry it at net 10,203,878 (EUR). **EXCLUDE** from every live month, whatever date the extract shows. |
| 7X000855 (JFC4, acompte, paid 31/08/2026) | 48,802,662.81 | Aug-26, frozen, as Fertilizers & Chemicals / Procurement | EXCLUDE from live months. Hardcoded today on File 26 `AJ784:AK784`. |
| 7X000851 (JFC2, acompte, pièce 1700000164, réf 2000001446, SAP date 2025) | 52,345,059.84 | Aug-26, frozen, as F&C / Procurement | EXCLUDE from live months. Hardcoded on File 20 `AJ108:AL108`; **confirmed by the user 23/10/2026** ("keep as is"). Aug-26 includes it. |
| 90021244 | 268,221 | Jul-26, frozen, as F&C | EXCLUDE from live months. Hardcoded on File 19 `AJ448`; **confirmed by the user 23/10/2026**. Jul-26 = 810,844,430 includes it. |
| 70000859 (ID DOC 221389, JESA, F&C / Procurement, advance; réf 2000001463, pièce 1700000068) | 408,659.37 | **Aug-26, frozen** (paid 24/08/2026) | **User decision 23/10/2026: paid 24 August → Aug-26.** The extract re-dated it to 08/09/2026, so every live extract shows it in Sep: EXCLUDE from live months. In the 19/10 file it was counted in neither month (File 20 `AL1920` and File 26 `AL2521` both blank); the frozen Aug-26 values must be corrected once when the slim workbook is built — see §F. |

## B. Pay Month forced (Scope = LIVE)

| 70001242 | 9,066,960.00 | SAP re-dated 18/05/2026 → 28/09/2026 (23/11 extract) | EXCLUDE (counted in frozen May-26) | Re-date sweep 23/11/2026; user to confirm. |
| 70001372, 70001371, 70001368, 70001367, 70001358, 70001352, 70000835 | 8,075,518.83 together | SAP re-dated 07/09/2026 → 31/08/2026 (23/11 extract) | HELD in Sep-26 (26/09/2026) pending user decision | Reported in Sep-26 since September. Moving them to Aug-26 means correcting the frozen Aug values (+8,075,518.83, F&C and W&E). |


| Invoice | Amount | Extract date | Pay Month | Reason |
|---------|--------|--------------|-----------|--------|
| 72000109 (JESA Buildings & Infrastructures advance, Water & Environement / Procurement) | 42,140,330.41 | 18/04/2025 | **26/10/2026** | Actually paid 01/10/2026; user instruction 12/10/2026. Today hardcoded on File 26 `AL5461`. |
| 70000649 (ID DOC 162944) | 2,254,861.76 | 07/09/2026 in every SAP extract to 23/11; the team's completed 23/11 file says 23/11/2026 | HELD in Sep-26 pending user confirmation (override 23/11/2026) | Reported in Sep-26 since September. |

## C. BU / Class assignments by the user (Scope = LIVE)

All hardcoded today on File 26 `AJ:AK` (yellow). Keep them as Overrides even where the AR
extract now resolves them, so a Mapping rebuild can never drop them.

| BU / Class | Invoices |
|------------|----------|
| Advisory / Proservices | 93000991 (3,988,796) |
| AM / Procurement (Sep-26, 21/09 decisions) | 90021934 (4,196,301), 90021783, 90021784, 90021785, 90021796, 90021798 |
| AM / Procurement (JESA invoice on JFC2, 9,063,016.82, paid 28/10/2026; decision of 02/11/2026) | 90022409 |
| AM / Procurement (JFC1, 22,123.50, paid 16/11/2026; decision of 16/11/2026) | 90022698 |
| Fertilizers & Chemicals / Procurement (Facility Management JORF, paid 11/11/2026; decision of 16/11/2026) | 90022483, 90022484, 90022485 |
| AM / Procurement (JESA invoices on JPH / MP1 / JFC1 / JFC2 / 107 DEF with no AR match; decisions of 06/10, 12/10 and 19/10) | 90022201, 90022202, 90022203, 90022204, 90022205, 90022206, 90022207, 90022209, 90022212, 90022216, 90022217, 90022218, 90022220, 90022224, 90022227, 90022234, 90022243, 90022244, 90022246, 90022290, 90022294, 90022325, 90022354, 90022360, 90022364, 90022365, 90022368, 90022376, 90022396, 90022403, 90022404, 90022425 |

Pattern for the future: unmapped JESA invoices on Division JPH / MP1 / JFC1 / JFC2 have **always**
been assigned AM / Procurement by the user. Claude still does not auto-assign: list them in
`Unmapped Invoices` with the suggestion "AM / Procurement (pattern)" and wait.

## D. Classification rules (not per-invoice)

0. **From Pay Month Nov-26 onward (user rule 16/11/2026, `Config!ClassRuleFrom`):** BU = AM keeps the
   AR `PSR/PROC` flag rule below. Every other BU uses AR col AJ `TYPE`: a value starting with
   "Labor" (not "Non") → Proservices; any other non-blank value ("Non Labor", "Procurement…",
   "Advance", "CC", "FC", "Subcontract") → Procurement; blank or "Labor and Non labor" → fall back
   to the flag rule and note it. `Mapping` col G carries TYPE. Sep-26 and Oct-26 keep rule 1.
1. Class comes from the AR `PSR/PROC` flag: `PSR` → Proservices, `PROC` → Procurement.
2. Flag precedence: Sheet3/Mapping col BH first, then the old `PSR-PROC Map`, then the
   fallback: Type 2 = "Services" → Proservices; Type 2 non-blank → Procurement; Type 2 blank and
   Invoice Type = "Labor" → Proservices, else Procurement.
3. The 7 AM invoices flipped PSR → PROC on 09/09/2026 by the user's BH edits (90020820,
   90020822, 90020937, 90021123, 90021386, 90021450, 90021572) are in frozen months; nothing to do.
4. Pay Month = 26th of the Date Paiement month. Payments dated 31/07 belong to Jul-26
   (the `Q2` input "Aug-26 window starts 31/07/2026" on the Summary is a leftover and is not
   used by any formula).
5. Jul-26 and Aug-26 are **never** repointed to a newer extract: a newer extract would replace
   91000988's 110.5M with 10.2M and understate Aug by ~100M.

## D2. Class overrides from the TYPE rule (Nov-26)

92004771 (W&E), 92004762 and 92004759 (Building & Infrastructure), 90022311 (F&C): AR TYPE = Non Labor
while the flag is PSR → Procurement. Entered as Overrides on 23/11/2026 because the Mapping TYPE column
was lost with the working copy; redundant once Mapping is rebuilt from an AR file with col G.

## Operational note (23/11/2026)

The working copy is lost whenever the cloud session's scratch space is cleared. The Drive copy is the
only durable state: it must be overwritten after every run. On 23/11 the Drive copy was still the
02/11 version and the 02/11–17/11 state had to be re-entered from this file and the session record.

## E. Reconciliation checkpoints (19/10/2026 baseline)

| Month | Gross in extract | Less overrides | Mapped = Summary | Proservices | Procurement |
|-------|------------------|----------------|------------------|-------------|-------------|
| Sep-26 | 257 inv / 407,774,514.59 | −70000859 (408,659.37) | 256 / 407,365,855.22 | 64 / 66,510,198 | 192 / 340,855,657 |
| Oct-26 | 149 inv / 174,915,937.60 | +72000109 (42,140,330.41) | 150 / 217,056,268.01 | 28 / 66,066,239 | 122 / 150,990,029 |

Grand total N16 = 7,464,017,093 in the 19/10 file; **7,464,425,753 after the §F correction.** The first
run must reproduce the Sep-26 and Oct-26 rows above from the same extract before computing anything
new, then report deltas.

## F. One-time correction to the frozen Aug-26 values (decision of 23/10/2026)

Add invoice 70000859 (408,659.37, Fertilizers & Chemicals, Procurement, 1 invoice) to Aug-26
when the slim workbook is built. Cells and resulting values:

| Cell | Content | 19/10 value | Corrected value |
|------|---------|-------------|-----------------|
| `K6` | Block A, F&C, Aug-26 | 246,268,477 | 246,677,136 |
| `K16` | Block A TOTAL Aug-26 | 855,789,161 | 856,197,821 |
| `J56` | Block D (PROC), F&C, Aug-26 | 239,959,708 | 240,368,367 |
| `J65` | Block D TOTAL Aug-26 | 719,484,723 | 719,893,382 |
| `F90` / `G90` | Block F, F&C, Aug-26 # / MAD | 57 / 239,959,708 | 58 / 240,368,367 |
| `F99` / `G99` | Block F TOTAL Aug-26 # / MAD | 268 / 719,484,723 | 269 / 719,893,382 |
| `C6` | # invoices F&C (all months) | 388 | 389 |
| `C16` | # invoices TOTAL | 2928 | 2929 |

Row totals (`N6`, `N16`, `M56`, `M65`) and `%` columns follow by formula. Block C (PSR) is
unchanged. Log the correction in `Data Sources`.

