# Slim workbook spec — `OCP Payments Weekly.xlsx`

Target size: 2–3 MB (the full report is 28 MB because `Sheet3` is a 37k × 60 AR extract
and `JESA OCP File 26` is an 11k × 42 paste). The slim file keeps only what the Summary
needs. The weekly SAP/OCP payment extract is supplied as a **separate file** every week
and is never pasted into the workbook.

## Tabs

| # | Tab | Purpose | Who writes it |
|---|-----|---------|---------------|
| 1 | `Summary by BU` | The report. Same layout as today (see below). | Claude, live-month cells only |
| 2 | `Config` | Frozen-through month, live months, header date convention, file names. | User (Claude reads) |
| 3 | `Overrides` | Every manual decision, one row per invoice. Replaces the yellow hardcodes in `AJ:AL`. | Claude on user instruction only |
| 4 | `Mapping` | Invoice → BU / Invoice Type / Type 2 / PSR-PROC. 5 columns, deduped. | Claude, only when a new AR extract is supplied |
| 5 | `Payments Key` | Key columns of the **last processed** extract, for the re-date sweep. | Claude, overwritten weekly |
| 6 | `Unmapped Invoices` | Live-month payments with no BU. Append-only. | Claude |
| 7 | `Data Sources` | One row per weekly run (audit log). Replaces the ever-growing note in `'Summary by BU'!B2`. | Claude |

Dropped from the slim file: `Sheet3` (replaced by `Mapping`), `PSR-PROC Map` (merged into
`Mapping` col E), `JESA OCP File NN` tabs (extract stays external), and all 20 external links
(`[1]OCP Payments October.xlsx`, AR files, P2P files). Everything those links fed is frozen as values.

## Tab 1 — `Summary by BU` (layout preserved)

Row/column positions are identical to the current file so the report looks the same.

| Block | Rows | Columns | Content |
|-------|------|---------|---------|
| A. Payments by BU | 3 header, 4–14 BUs, 16 TOTAL | B BU · C # invoices · D:I Jan–Jun · J Jul · K Aug · L Sep · M Oct · N Total · O % | Net à payer by BU and Pay Month |
| B. Non-OCP & subsidiary | 18–19 titles, 21 header, 22–30 BUs, 31 total | C:F Jan–Apr · G Total · H % | Static (AR tab, Project Client ≠ OCP). Not refreshed weekly. |
| C. PSR all months | 35 title, 37 header, 38–48 BUs, 49 TOTAL | C:H Jan–Jun · I Jul · J Aug · K Sep · L Oct · M Total | Proservices by BU |
| D. PROC all months | 51 title, 53 header, 54–64 BUs, 65 TOTAL | same as C | Procurement by BU |
| E. PSR Jul–Oct detail | 68 title, 69 source note, 70 header, 71–81 BUs (`=B4..=B14`), 82 TOTAL | C,D,E = Jul (# · MAD · %) · F,G,H Aug · I,J,K Sep · L,M,N Oct | counts + amounts + share |
| F. PROC Jul–Oct detail | 85 title, 86 source note, 87 header, 88–98 BUs, 99 TOTAL | same as E | counts + amounts + share |

BU order (rows 4–14, reused by `=B4..=B14` in blocks E/F): Water & Environement, AM, … ,
Building & Capabilities (11 rows). Copy the names exactly — SUMIFS/lookups are case- and
accent-sensitive ("Environement" is spelled that way in the data).

Month headers: `D3:I3` are text (`Jan-26` … `Jun-26`); `J3` onward are **dates = 26th of the
month** (26/07/2026, 26/08/2026 …). The 26th is the Pay Month key. Keep that convention when
adding a column.

### Frozen vs live

* **Frozen** (paste as values, never recompute): every column up to and including the month in
  `Config!FrozenThrough` (today: Aug-26). In the current file these are
  `D4:K16`, `C22:H31`, `C38:J49`, `C54:J65`, `C71:H82`, `C88:H99`. Their values currently come
  from the external workbook (`JESA OCP File 19/20`, `File 14` ratios) and from cached cells —
  once frozen they are the audited history. Jul-26 = 810,844,430 and Aug-26 = 855,789,161
  (see standing-exceptions.md for what those two numbers contain).
* **Live**: Sep-26 and Oct-26 today (`L4:M16`, `K38:L49`, `K54:L65`, `I71:N82`, `I88:N99`),
  recomputed from the extract every week. Totals (`N`, `M`, `SUM` rows) and `%` columns stay as
  formulas.
* When the extract brings a payment dated in a **new** month: insert a column before `Total`
  in blocks A, C, D (copy the previous month's column formatting, header = 26th of the new
  month) and add a 3-column group (#, MAD, %) before the end of blocks E and F. Extend the
  TOTAL and `%` formulas. Report the insertion.
* When the user says a month is closed: set `Config!FrozenThrough`, paste that month's live
  cells as values, log it in `Data Sources`.

Live cells may be written either as values computed in Python (default — robust with
openpyxl, no recalc needed) or as `SUMIFS` against a `Payments` tab if the user wants formulas
back. The reconciliation step must hold either way.

### Cell `B2`

Keep one short line: `Source: Jan–Aug-26 frozen values (see Data Sources). Live months
from weekly SAP/OCP extract via Mapping + Overrides.` Do **not** append run notes here; they
go to `Data Sources`.

## Tab 2 — `Config`

| Key | Value (today) |
|-----|---------------|
| FrozenThrough | 2026-08-26 |
| LiveMonths | 2026-09-26, 2026-10-26 |
| PayMonthDay | 26 |
| ExtractLayoutHash | hash of the 35 header names A1:AI1 of the last accepted extract |
| MappingSource | file name + date of the AR extract used to build `Mapping` |
| LastExtract | file name, max Date Paiement, row count |

## Tab 3 — `Overrides`

| Col | Field | Notes |
|-----|-------|-------|
| A | Invoice (Réference) | as text, trimmed, no NBSP |
| B | BU | blank = keep lookup |
| C | Class | Proservices / Procurement / blank |
| D | Pay Month | date (26th) · `EXCLUDE` · blank = keep extract date |
| E | Scope | `LIVE` (apply to live months) · `FROZEN` (already in a frozen month — must never reappear live) |
| F | Reason | one line |
| G | Decided by / on | user + date |

Precedence: Overrides → Mapping → fallback rules. Seed rows: see standing-exceptions.md.

## Tab 4 — `Mapping`

Built from the Power BI AR extract (`AR` tab): A Invoice Number (col P) · B Business Unit
(col O) · C Invoice Type (col AL) · D Type 2 (col AW) · E PSR/PROC (col BH). Trim and strip
NBSP from invoice numbers; dedupe on invoice; if two rows disagree on BU or flag, keep the
row with the latest Payment Date (col AD) and list the conflict in the run report. Today's
source is `Sheet3` (37,081 rows, max Payment Date 01/10/2026) → about 27–37k unique rows.
Rebuild only when a new AR extract is supplied; never hand-edit (decisions go to Overrides).

## Tab 5 — `Payments Key`

Last processed extract, paid rows only (`Net à payer ≠ 0`): Réference (I) · ID DOC (C) ·
Fournisseur (G) · Division (U) · Pièce comptable (AC) · Date Paiement (AD) · Réf Paiement (AE) ·
Net à payer (Z) · Pay Month · BU · Class · Extract date. ~1,200 rows. Used to detect re-dated
payments (same AC + AE, different month) and to produce week-over-week deltas by invoice.

## Tab 6 — `Unmapped Invoices`

Columns A:L as today: Extract row · Réference · ID DOC · Fournisseur · Service · Division ·
Date Doc · Date Paiement · Pay Month · Réf Paiement · Net à payer · BU to assign (left blank
for the user). Append a dated block each run; mark rows resolved when an Override lands.

## Tab 7 — `Data Sources`

Columns as today: Written to · Source · Object · Obtained via · Query or request · Parameters ·
Run at · Requested by · Row count · Changes after retrieval · Notes. One row per run plus one
row per Mapping rebuild / freeze / override batch. Mark superseded rows in Notes.

## Weekly extract file (input, not stored)

`JESA OCP File NN` / SAP supplier-payment export, 35 columns `A:AI`, header in row 1:

| Col | Header | Use |
|-----|--------|-----|
| C | ID DOC | identity, Unmapped list |
| G | Fournisseur | Unmapped list |
| I | Réference | invoice number → Mapping / Overrides key |
| U | Division | Unmapped list (JPH / MP1 / JFC1 / JFC2 pattern) |
| Z | Net à payer | amount summed |
| AC | Pièce comptable | payment identity (with AE) |
| AD | Date Paiement | Pay Month = DATE(year, month, 26) |
| AE | Réf Paiement | payment identity (with AC) |

Full header list for the layout check: Code Sté, Sté, ID DOC, Site Scan, Service, Code Fsseur,
Fournisseur, Type Doc, Réference, Date Doc, Montant Doc, Devise, Contrevaleur MAD, Date base
échéance, Cdp, Date Echeance, Echeance, Commande, Contrat, Type Commande, Division, Statut,
Rôle Actuel, Desc Statut, Code Blocage, Net à payer, Date création, Facture MM, Pièce comptable,
Date Paiement, Réf Paiement, Desc Statut1, Date dernier Commentaire, Dernier Commentaire, Filiale.
Any difference → stop and report before computing anything.
