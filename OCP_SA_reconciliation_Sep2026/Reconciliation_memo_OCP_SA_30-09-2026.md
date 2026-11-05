# JESA / OCP SA intercompany reconciliation at 30 September 2026

**Purpose.** OCP SA's reconciliation team reports a net claim on JESA of **4,588.2 M** (OCP supplier accounts 3411 / 4411 / 4413 / 4417 / 4481). JESA's books show a net amount owed to OCP SA of **447.5 M** (customer 4100410, accounts 3421 / 3423 / 3427 / 4421 / 4491 and 4411 / 4417). The difference of **4,140.7 M** is explained below line by line, with the action required on each side. Detail is in the workbook `JESA_OCP_SA_reconciliation_30-09-2026.xlsx` (tabs Summary, By project, By PO, and one tab per reconciling class).

Sign convention used everywhere: **positive = OCP owes JESA**, negative = JESA owes OCP. All figures in MAD, TTC, as at 30/09/2026.

## 1. Positions per ledger

| JESA books (customer OCP SA) | MAD | OCP SA books (supplier JESA) | MAD |
|---|---:|---|---:|
| 3421 Invoices issued, open | 1,092.5 M | 3411 Advances / down-payments paid (debit) | 7,009.6 M |
| 3423 Retentions (RG) receivable | 162.6 M | 4411 Invoices booked, open | -345.6 M |
| 3427 Unbilled revenue (POC) | 4,497.7 M | 4413 Retentions | -8.6 M |
| 4421 Advances received (net of recoveries) | -5,790.9 M | 4417 Receptions not invoiced (GR/IR) | -2,066.8 M |
| 4491 Deferred revenue | -400.9 M | 4481 Fixed-asset suppliers | -0.4 M |
| 4411/4417 OCP as supplier of JESA | -8.5 M | | |
| **JESA net (OCP owes JESA)** | **-447.5 M** | **OCP net restated in the same sign** | **-4,588.2 M** |

**Difference to explain: 4,140.7 M.**

## 2. What the difference is made of

The two ledgers were matched at line level (invoice numbers, advance-invoice numbers, PO numbers, legacy PO numbers and work-order codes). Every OCP line and every JESA line was put in one of the classes below; the classes add up exactly to the difference.

| Class | Driver | MAD | Who acts |
|---|---|---:|---|
| **Billing gap** (JESA 3421+3423 vs OCP 4411+4413+4481) | | **900.6 M** | |
| B1 | JESA invoices open at JESA with no trace in OCP's ledger (847 invoices, of which 326.9 M dated 2026) | 542.1 M | OCP: record / release in e-invoicing and SES |
| B2 | JESA advance invoices issued and not yet paid by OCP (gross-up at JESA, zero net) | 280.0 M | none (presentation) |
| B3 | Invoices open on both sides: residual differences (RG, partial payments, VAT withholding) | -31.0 M | both |
| B4 | Invoices booked and settled per OCP's ledger, still open at JESA (payment in transit / cash to apply) | 187.2 M | JESA (cash application) / OCP (payment advice) |
| B5 | OCP open payables (mostly REGUL lines) with no open item at JESA | -77.8 M | OCP: match REGUL lines with 3411 KA |
| **Unbilled gap** (JESA 3427+4491 vs OCP 4417) | | **2,029.9 M** | |
| U | JESA unbilled revenue net of deferred revenue | 4,096.7 M | info |
| U1 | OCP migrated GR/IR items dated 2016-2021 (Jacobs legacy), no JESA counterpart | -293.9 M | OCP: clean-up |
| U2 | OCP receptions (SES) accepted, not yet invoiced | -1,772.9 M | both: SES release / invoicing |
| **Advance gap** (JESA 4421 vs OCP 3411) | | **1,218.7 M** | |
| A1 | Advance invoices matched to OCP payments (by reference or identical amount), residual | 4.8 M | both |
| A1u | JESA advance invoices with no open advance left at OCP (unpaid, incl. B2, or already fully cleared by OCP) | -1,402.9 M | OCP: confirm payment / reference |
| A2a | OCP payments of JESA progress invoices carried by OCP as down-payments | 549.6 M | OCP: clear 3411 against the invoice |
| A2c | Advance invoices paid by OCP, fully recovered by JESA, still carried in full by OCP | 234.8 M | OCP: post the regularisations |
| A2b | OCP down-payment lines JESA cannot identify (unreferenced, pre-2022 legacy, AUX residuals) | 123.3 M | OCP: provide settled JESA document numbers |
| A3u | Advance recoveries made by JESA on progress invoices OCP has no record of | 2,229.1 M | OCP: process the invoices and post KA |
| A3f | Advance recoveries made by JESA on invoices OCP has booked, recovery not posted by OCP | 1,044.0 M | OCP: post KA regularisations |
| A4 | Advance recoveries / reversals already booked by OCP | -1,564.0 M | info |
| | JESA 4411/4417 (OCP as supplier of JESA, outside OCP's extract) | -8.5 M | |
| | **Total** | **4,140.7 M** | |

## 3. Reading of the difference

1. **OCP's claim is a cash position, not a receivable.** OCP has paid JESA 7,009.6 M of advances and down-payments and has accepted receptions (SES) of only 2,066.8 M plus 354.5 M of booked invoices. In JESA's books the same advances (5,790.9 M net) have largely been consumed by work performed: 4,497.7 M of unbilled revenue recognised on percentage of completion and 1,255.1 M of issued invoices.

2. **Advance mechanics drive 1,218.7 M of the gap.** JESA has recovered 3,273.0 M of advances through its progress invoices (the recovery is deducted on the invoice, so most of these invoices needed no cash and never show as "unpaid"). On OCP's side, 1,564.0 M of regularisations are posted and, in addition, advance invoices of 1,402.9 M are fully cleared in OCP's 3411 while their credit is still open in JESA's 4421 (A1u). Net of both, OCP's recognition of recoveries lags JESA's by 306.1 M. The point to note is that 2,229.1 M of JESA's recoveries sit on 1,448 progress invoices that have no trace at all in OCP's ledger: OCP has cleared its advances against other documents, so a document-by-document matching of the recoveries is required (tabs "A3u", "A3f", "A1u", one line per invoice).

3. **OCP's 3411 also contains 907.8 M that are not advances.** 549.6 M are payments of JESA progress invoices that OCP booked as acomptes ("OCP_E-INVOICE Invoice With PO") without ever recording the invoice; JESA applied that cash to the invoices, so there is nothing left to recover. 234.8 M are advance invoices JESA has fully recovered through progress invoices and OCP never regularised. 123.3 M are residual lines (unreferenced payments, pre-2022 legacy, AUX splits) OCP should identify.

4. **542.1 M of JESA invoices are not in OCP's books at all** (tab "B1"). 326.9 M are dated 2026 and are the normal e-invoicing / SES pipeline; the older ones (2021-2024) need to be re-submitted or confirmed as received by OCP. Conversely 187.2 M appear settled by OCP but are still open at JESA: OCP's payment advice is needed to apply the cash (payments made around 30/09 are likely in transit).

5. **The remaining 2,029.9 M is timing on revenue recognition.** JESA recognises revenue on percentage of completion; OCP accrues only on approved service entry sheets. The JESA AR register gives the status of every unbilled amount per project (pending client PM approval, signed by client PM pending OCP SS approval, pending client PO, SSS postponed...); it is reported per project in the workbook. Inside this bucket, OCP still carries 293.9 M of migrated 2016-2021 receptions (Jacobs legacy, POs 249, 273, 4430...) that have no JESA counterpart and should be cleaned; note that cleaning them *increases* OCP's net claim, so it does not help close the gap.

6. **Entity scope.** JESA bills part of the TSP / DAP / JFC programmes to OCP Nutricrops and JFC I-II-IV-V (separate customer accounts, not in this extract) while OCP SA's ledger carries the POs. 30 projects are flagged in the "By project" tab with the split of JESA's open positions by entity; both sides must align the counterparty before the balances can agree PO by PO.

7. **Data quality of the extracts.** OCP's 3411 extract lists 167 documents twice with opposite signs (1.29 bn gross, net zero); they were excluded from the matching. JESA's 3427 carries 21,700 migrated lines (account 3427900999) at WBS level without PO.

8. **Why a PO-by-PO comparison alone cannot work.** JESA's ledger is kept by WBS; the PO is tagged on part of the lines only and, for the SP2M / TSP / DAP programmes, as groups of POs (one advance covering 10 to 40 POs). OCP's migrated items carry legacy PO numbers (7xxx, 8xxx, 28537-10...) that were linked to the current POs through the advance-invoice references. The reconciliation is therefore built at project level (group of POs) with the PO detail underneath ("By PO" tab).

## 4. Largest projects (JESA net vs OCP net)

| Project | BU | Description | JESA net | OCP net | Difference | Billing gap | Unbilled gap | Advance gap |
|---|---|---|---:|---:|---:|---:|---:|---:|
| QC.01546 | SPM | CENTRAL AXIS-MPH OSBL | -71.6 M | -707.0 M | 635.3 M | 182.9 M | 272.3 M | 180.1 M |
| QC.01547 | SPM | SP2M PAP_Construction & Commssioning | 13.2 M | -586.8 M | 600.1 M | 44.0 M | 560.0 M | -3.9 M |
| QC.01717 | SPM | Safi Existing Site and ES-SPH Corridor | 72.5 M | -184.1 M | 256.6 M | 6.5 M | 215.4 M | 34.7 M |
| QC.01466 | FEC | New TSP Hub (2 units) | 206.6 M | -2.4 M | 209.0 M | 5.4 M | 125.4 M | 78.2 M |
| QC.01432 | WEN | Water Emergency Program | 137.2 M | -44.7 M | 181.8 M | 17.7 M | 74.0 M | 90.1 M |
| QC.01715 | SPM | Central Axis-MS | 43.0 M | -118.4 M | 161.3 M | 0.0 M | 152.7 M | 8.6 M |
| QC.01725 | SPM | SP2M - LOUTA MINE PMC Role | -376.8 M | -530.9 M | 154.2 M | -4.7 M | 139.4 M | 19.5 M |
| QC.01627 | SPM | Central Axis - NSP-SPH Platform | -102.7 M | -238.8 M | 136.0 M | 1.7 M | 56.7 M | 77.7 M |
| QC.01866 | FEC | Molten Sulfur Project U363 | 20.4 M | -106.5 M | 127.0 M | -12.3 M | -73.3 M | 212.5 M |
| QC.01548 | SPM | SP2M SAP - Construction | -9.3 M | -126.2 M | 116.9 M | 0.0 M | 86.6 M | 30.3 M |
| QC.01160 | AMT | Jorf Lasfar North Fertilizers Plant Y3 | 99.4 M | -15.9 M | 115.3 M | 10.8 M | 44.3 M | 60.2 M |
| QC.01746 | FEC | TSP Hub Handling and Storage C+F | 64.9 M | -45.3 M | 110.2 M | 0.0 M | 83.5 M | 26.7 M |
| QC.01483 | SPM | CENTRAL AXIS-MPH MASTER PLAN | 47.7 M | -54.7 M | 102.4 M | 0.0 M | 8.8 M | 93.6 M |
| QC.01694 | ENE | MPH Electrical project | 2.4 M | -88.7 M | 91.1 M | 5.3 M | 65.0 M | 20.7 M |
| QC.01683 | FEC | Global Cost + Fee for PAT//PPA/MAPc | -191.8 M | -282.2 M | 90.4 M | 2.4 M | 56.2 M | 31.7 M |

The per-project explanation sentence and the AR status of the unbilled amounts are in the "By project" tab; 29 projects explain 80% of the absolute difference.

## 5. Actions proposed

**OCP SA**
1. Record the 847 JESA invoices in tab "B1" (542.1 M) and match, invoice by invoice, the 1,448 recovery invoices in tab "A3u" against the regularisations OCP has posted (tabs "A1u" / "A4"), so that each advance invoice shows the same remaining balance on both sides.
2. Record the invoices behind the 549.6 M of 3411 items that settled progress invoices and clear them (tab "A2"), post the regularisations on the 234.8 M of fully recovered advances, and document the 123.3 M of residual lines.
3. Post the missing KA regularisations on invoices already booked (1,044.0 M, tab "A3f") and match the REGUL lines in 4411 with 3411 (77.8 M, tab "B5").
4. Clean the 2016-2021 migrated GR/IR (293.9 M, tab "U1") and provide payment advices for the invoices in tab "B4".

**JESA**
1. Apply the cash for the 187.2 M of invoices OCP shows as settled, once OCP provides the payment advices.
2. Re-submit the pre-2025 invoices in tab "B1" to OCP, and confirm with OCP the status of the advance invoices in tab "A1u" (1,402.9 M: unpaid, or fully regularised on OCP's side).
3. Complete the PO tagging in the GL (N° PO field) for the projects listed without PO, so that the next reconciliation can be run PO by PO.
4. Keep the unbilled position documented by status (AR register) as support for the POC revenue; this is the part of the gap that is expected and will reverse as SES are approved.

Neither ledger is "wrong" on the economics: OCP has paid in advance, JESA has performed and invoiced against those advances. The gap is almost entirely made of documents JESA has issued and OCP has not yet processed, plus the normal POC-vs-SES timing. No reserve is indicated on JESA's side on the basis of OCP's balance; the exposure to monitor is the ageing of the unbilled and of the invoices OCP has not yet recorded.
