# Workbook map — CASH FLOW ANALYSIS - BUs

Master file: `cashflow-bu/workbook/CASH_FLOW_ANALYSIS_-_BUs.xlsx` (59 tabs, ~11 MB).
Every figure on the report is a SUMIFS over a source tab. Nothing is typed twice.
Scope: JESA Group, 10 business units, MAD. Cash OUT is stored **negative**, cash IN **positive**.

## 1. Output tabs (read these, never type into them)

| Tab | Role |
|---|---|
| `EXEC Summary` | Cover pack. Controls (E6 master switch, E7 data-as-of), headline row 13, sections 1–5. |
| `EXEC - BU & Integrity checks` | Net position by BU + 18 integrity checks in rows 23–40 (col D difference, col F status, all must read `OK`). |
| `REPORT - Cashflow by BU` | The detail report. Columns B/C/D = Jul/Aug/Sep actuals (row 4 holds the month start date), E = October actuals, F = Oct ongoing approved, G = net after approved, H = Oct ongoing pending approval, I = net after all ongoing, K = BU key (auto). P2 = master switch (1/0). |
| `HOW TO UPDATE` (hidden) | Built-in runbook (six refresh steps, data flow, input register, month roll-forward). |
| `Data Sources` | Audit trail — one row per load (A written-to, B source file, C object, D obtained via, E request, F parameters, G run at, H requested by, I row count, J changes after retrieval, K notes). 101 rows at W3. Append, never edit. |
| `AM - Cash IN Invoices`, `AM - Cash OUT Invoices`, `OPS - Cash In (Invoices)`, `OPS - Cash Out (Invoices)` | Invoice-level views for AM and the operating BUs (derived). |

### BU blocks on `REPORT - Cashflow by BU`
Consolidated block rows 7–37 (`JESA TOTAL`). BU block header rows: AM 40 · Advisory 72 · Building & Capabilities 101 · Building & Infrastructure 130 · Energy Operations 159 · Fertilizers & Chemicals 188 · Global Mining 217 · JESA Technologies 246 · SP2M Ad Hoc Program 275 · Water & Environement 304.
Blocks share the same line labels but not the same length, so locate lines by the label in column A, never by offset: AM has 30 lines (extra `Less PSR (Labor)` and `Less PISTELLI / PELZ` rows), Water has 28 (extra `Less Overhead (OH)` row), the other eight have 27. Line order: header · PSR · Revenue — PSR · PSR (Labor) · [Less PSR] · PSR reserve provision · NET PSR · PROCUREMENT · Revenue — Procurement (PROC) · Procurement (Non-Labor) · Cash IN — PROC Compensable (standing) · Unclassified collections · Total Cash IN · tiers `1` `2` `3` `Intl 1` `Intl 2` `Intl 3` `Etranger*` `OCP et Filiale` `Office/Publique` `Unclassified / Not on OCP` · LC payments · Customs · [Less Overhead (OH)] · [Less PISTELLI / PELZ] · Total Cash OUT (incl. LC & customs) · NET BALANCE · of which Overhead (OH).
Column J on the LC row carries the project code used to filter the LC source: AM=`AMT`, ENE=`ENE`, F&C=`FEC`, JET=`JET`, SP2M=`SP2M`, Water=`WEN` (blank = no LC for that BU).

## 2. Source tabs → what they feed (the formula criteria)

### Cash IN — `SRC - Cash In (AX Jul-Aug)` (OCP payment report, per invoice)
Despite the name it holds Jul → Oct. Rows 2:30086 are in the SUMIFS ranges.
Raw columns A:AI come from the OCP payment report (AX export): `C ID DOC`, `G Fournisseur`, `I Réference` (invoice), `J Date Doc`, `K Montant Doc`, `M Contrevaleur MAD`, `R Commande`, `U Division`, `Z Net à payer` (**amount used**), `AC Pièce comptable`, `AD Date Paiement` (**date used**), `AE Réf Paiement`, `AI Filiale`.
Helper columns the formulas key on: `AO BU Mapped` (canonical BU name or `Unmapped`), `AP Labor/Non-Labor` (`Labor` → PSR rows, `Non-Labor` → Procurement rows), `AN PSR/PROC (AR match)`, `AQ Type 2 (AR)`, `AR Project / WBS (AR)`, `AS AM OPEX-CAPEX`, `AT Update note`.
Report formula pattern: `SUMIFS(Z, AO, <BU>, AP, "Labor"|"Non-Labor", AD, >=month start, AD, <=month end)`. Unclassified collections = all mapped collections in the month − PSR − PROC.
Dedupe key when loading: invoice reference (I) + payment reference (AE) + amount (Z); the OCP file is cumulative so most rows are already present.

### Cash OUT actuals — SAP payment extracts (one tab per month)
| Month | Tab | Amount | BU | Tier | OH flag | Include flag |
|---|---|---|---|---|---|---|
| Jul | `SRC - Payments Jul (Mapped)` | L `Amount in MAD` (negative) | A `BU` | B `Tiering` | O `OH?` | — |
| Aug | `SRC - Payments Aug (Mapped)` | L | A | B | O | — |
| Sep | `September Actual Payments` | C `Montant en MAD` | AK `BU (clean)` | AM `Tier (report)` | AP `OH?` | — |
| Oct | `October Actual Payments` | C `Montant en MAD` (positive, report negates) | AH `BU (mapped)` | AG `Tier (report)` | AJ `OH?` | AF `Include in Oct cash OUT?` must be exactly `Yes` |
Sept/Oct raw columns A:AE are the SAP FBL1N-style extract: `A New Tier`, `B Type`, `C Montant en MAD`, `D Paiement positive` (Oct), `F/G Supplier/Name`, `H Document Number`, `I Document Type`, `P Reference`, `Q Clearing Document`, `R Clearing Date`, `W Posting Date`, `AC Purchasing Document`.
`AF` values seen: `Yes`, `No - invoice line`, `No - corporate OH`, `No - excluded tier` (Agency Local / Filiale), `No - payment reversed (net zero on clearing doc)`, `No - split by BU (rows …)`.
`AH` = `(unmapped)` when no BU could be found; the consolidated block picks unmapped lines up in the `Unclassified / Not on OCP` row, BU blocks do not.
`AI BU mapping source` records how the BU was found (ongoing-position match, user BU on invoice list, supplier history). `AK Matched ongoing row` links the payment to the OctOngoing row it closed.
Dedupe key: Document Number + Clearing Document + amount.
**Roll-forward:** a new month needs a new tab (copy the October tab layout) and a new report column — see `HOW TO UPDATE` section D.

### Cash OUT ongoing — `REPORT - Oct Ongoing-Position`, Excel table `OctOngoing` (A1:AH1362 at W3)
One row per open commitment (invoice / attachment / balance). New rows must be inserted **inside** the table.
Columns the report keys on:
- `A BU` (canonical name), `B Supplier`, `C PO`, `D Reference / Invoice`, `E Amount (MAD)` (**negative**), `F Due date`, `G Tier` (`1`,`2`,`3`,`Intl 1`,`Intl 2`,`Intl 3`,`Unclassified`,`Not on OCP Tier`), `H Priority/Status`, `I Comment`, `J Source`.
- `N Sept payment status` — the **lifecycle status**. Counted only when it begins with `Open` (`Open`, `Open - AM approved list W1 (12/10)`, `Open - WEN Batch 5 (19/10)` …). Closed values: `Paid in Sept`, `Paid in Oct`, `Deferred - due after Sept`, `Superseded - …`, `Excluded - …`, `Removed - …`.
- `Q Manual status override` — when set, N follows Q. This is where a line is closed (`Paid in Oct`) or reinstated.
- `P OH?` = `OH` for overhead suppliers (looked up on `OH Suppliers`); OH lines are excluded from the procurement headline via the "Less OH" rows.
- `U In payment tracking log? (excl. corporate)` Yes/No, `V/W` log status and detail.
- `AA Approved payments (Oct)? — in tracking log or on Priority list 02.10` — **Yes = approved (report col F), No = pending approval (report col H)**.
- `AB Oct position` text, `AC/AD` 09/10 log check, `AG/AH` 20/10 (W3) log check, `Y Date added`, `Z Date added — basis`, `AE/AF` FX note (EUR→MAD 10.8 ex-M1, 10.66 F&C Sept list).
Report pattern: approved = `SUMIFS(E, A, <BU>, G, <tier>, N, "Open*", AA, "Yes")`; pending = same with `AA = "No"`. AM `PISTELLI*` lines are excluded through row 66 of the AM block.
Dedupe key: normalised reference (`O Ref key (norm)`) + supplier; amount-only matches are reviewed on `Review - Amount-only matches`.

### LC payments — `SRC - LC Payments` (bank statement lines) and `LC Ongoing (19-10)` (bank LC situation)
- Actuals: `SUMIFS('SRC - LC Payments'!G Débit(DH), S Project, <code J>, E Date in month)`. Columns: `E Date`, `F Libellé`, `G Débit(DH)` (positive), `I MNT DEVISE`, `J D.Curr`, `K PO`, `M Ref LC`, `N Nom fournisseurs`, `S Project` (AMT/ENE/FEC/JET/SP2M/WEN/PAP…), `X Vendor`, `Y BU`, `AE Month`, `AF Tier`, `AG Origin`.
- Ongoing (approved) = `SUMIFS('LC Ongoing (19-10)'!Q, P, <BU>)`: `P Business Unit (report)`, `Q Oct ongoing (MAD, signed)` (negative; 0 when `R Already paid in Oct? = Yes`), `L Oct-26 LC maturity (MAD)`, `C Project code`, `D PO`, `E LC / CREDOC ref`, `F Supplier`, `H Ccy`, `J Encours LC (FX)`, `K Encours LC (MAD)`. Range is pinned to rows 5:32 — a new LC situation means a new tab (`LC Ongoing (dd-mm)`) **and** repointing the LC-row formulas in column F of all blocks (rows 63, 181, 210, 268, 297, 326 and the consolidated row 30 via its sum).
- Known data issues: row 47 is column-shifted and deliberately left broken (duplicate of row 28); the integrity tripwire counts non-numeric amounts.
- Dedupe key: Date + Débit + Ref LC.

### Customs — `DD` (Treasury customs detail, current method since 12/10)
`SUMIFS(DD!E Montant sur relevé, I BU (report), <BU>, J Month, <month start>)`, negated. Columns: `B Date`, `C Valeur`, `D Libellé d'opération` (e.g. `PAIEMENT TGR HB 02/07/2026`), `E Montant sur relevé`, `F Montant du ticket`, `H BU` (code), `I BU (report)` = XLOOKUP of H on the code map `L2:M11` (AMT→AM, FEC→Fertilizers & Chemicals, SP2M, WEN→Water & Environement, JET→JESA Technologies, ENE→Energy Operations, GMM→Global Mining, BIN→Building & Infrastructure, ADS→Advisory, BIC→Building & Capabilities), `J Month` = first day of month of B (formula). Ranges run to row 2000.
Older tabs `SRC - Custom Fees`, `Customs Pmts (RECEVEUR DOUANE)`, `LC & Customs (Jul-Sep)` are history only.
Dedupe key: Date + amount + BU code.

### Standing manual inputs (the only typed numbers)
| Input | Where | Note |
|---|---|---|
| Master switch | `EXEC Summary`!E6 (TRUE/FALSE) | mirrored to report P2 |
| Data as of | `EXEC Summary`!E7 | the date you write on the cover |
| PSR reserve | `PSR Reserve (provision)`!B5 total, B6 months | spread to Water rows by formula |
| PROC compensable (dividends) | `PROC Compensable (standing)`!C5:C14, GM invoices A22:C26 | standing, overwrite only with a new schedule |
| Monthly revenue (P&L) | `SRC - Monthly Revenue (<month>)` F7/F8 (group) and per-BU columns row 7/8 (MMAD) | one tab per month; report multiplies by 10^6 |
| SP2M Sept customs | report D254 | historical plug, now sourced from DD |

### Reference / mapping tabs
`Tiering Master` (supplier → tier, OCP database first), `OH Suppliers` (overhead supplier list, drives every `OH?` column), `PO mapping ` (PO → WBS, 13k rows), `SRC - SBU Cliente (AR)` (project → OCP MD), `AM OPEX-CAPEX`, `AM OPEX projects`, `Priority list (02.10)` (approved priority invoices; flagged `AA = Yes` on OctOngoing), `SRC - Invoice Tracker (02.10)` / `(24.09)` (snapshots of the Payment tracking log — Invoice Tracker table: `A Request date`, `B BU`, `D Supplier`, `G Invoice Reference`, `I/J/K` amounts, `O Accounting status`, `R Type de blocage`, `S Treasury Comments`, `U Payment Date`, `V Overall status`).

## 3. Canonical lists
BUs (exact spelling, used as SUMIFS criteria): `AM`, `Advisory`, `Building & Capabilities`, `Building & Infrastructure`, `Energy Operations`, `Fertilizers & Chemicals`, `Global Mining`, `JESA Technologies`, `SP2M Ad Hoc Program`, `Water & Environement` (single n — keep it).
Codes: AMT, ADS, BIC, BIN, ENE, FEC, GMM, JET, SP2M (also SP2M-PAP, SP2M-OSBL, PAP, MPH OSBL → SP2M), WEN.
Tiers on the report: `1`, `2`, `3`, `Intl 1`, `Intl 2`, `Intl 3`, `Etranger*`, `OCP et Filiale`, `Office/Publique`, `Unclassified / Not on OCP` (source values `Unclassified`, `Not on OCP Tier`). Excluded tiers: `Agency Local`, `Filiale`.
Headline (EXEC row 13): C collections OCP report · D dividends · E cash out Jul–Sep · F cash out Oct actual · G approved payments · H net after approved · I pending payments · J net position all ongoing. PSR line row 16.

## 4. Excel features to respect when scripting
- `OctOngoing` is a real table: extend `ws.tables['OctOngoing'].ref` when appending rows.
- Since 02/11/2026 the formulas that feed the report (REPORT PSR reserve / PROC compensable rows, DD!I, LC Ongoing!P, September AL, OctOngoing G tier) are INDEX/MATCH, so LibreOffice evaluates them and `deliver.py` ships a file with calculated values. The derived view tabs (AM/OPS invoice views, Synthèse, Détail par fournisseur, bridges) still use FILTER/UNIQUE/LET/VSTACK: LibreOffice leaves `#NAME?` there and Excel recomputes them on open (full calc on load is set).
- Helper columns that are **formulas** (copy them down, never type values): OctOngoing N status, O ref key, P OH?, AA approved, AB position · `September Actual Payments` AK/AL/AM/AN/AO/AP · `SRC - Payments Jul/Aug` N/O/P · `SRC - LC Payments` AF/AG · `LC Ongoing (dd-mm)` P/Q/R · DD I/J · `SRC - Cash In` AS. Helper columns that are **values**: `October Actual Payments` AF:AK, `SRC - Cash In` AN:AR/AT, `SRC - LC Payments` AE.
- Approval on OctOngoing is driven by `U In payment tracking log?` = Yes (or the reference on `Priority list (02.10)`); user-approved lists are recorded as U = Yes with the basis in X ("treated as in tracking log per user").
- The master switch cell is `EXEC Summary`!F6 (E6 is empty); the data-as-of date sits in E7 and F7.
- Never load with `data_only=True` and save — formulas would be lost.
- Keep the header rows; SUMIFS ranges run to row 20 000 (payments), 30 086 (cash in), 2 000 (DD, October payments). Extend ranges if a tab grows past them.
