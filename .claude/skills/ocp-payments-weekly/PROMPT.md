# Weekly prompt (paste as-is, attach the two files)

```
/ocp-payments-weekly

Weekly OCP payments refresh.
Attached:
  1. OCP Payments Weekly.xlsx   (slim workbook: Summary by BU + Config + Overrides + Mapping + Payments Key + Unmapped Invoices + Data Sources)
  2. JESA OCP File <NN>.xlsx    (this week's SAP/OCP payment extract, 35 columns A:AI)
  3. [optional] Power BI - AR <date>.xlsx  (only if the mapping must be rebuilt)

Rules:
- Frozen months (Config!FrozenThrough and earlier) are never touched. Never repoint Jul-26 / Aug-26.
- Apply every Overrides row before any lookup; Overrides > Mapping > fallback rules.
- Invoices 91000988, 7X000855, 7X000851, 90021244 are in frozen months: exclude them from live months even if the extract re-dates them.
- Do not assign a BU yourself. Unmapped live-month payments go to Unmapped Invoices; I will answer and you write my answer to Overrides.
- Re-date sweep against Payments Key (same Pièce comptable + Réf Paiement, different month): report, do not fix.
- Reconcile: gross extract per live month → ± overrides → mapped total = Summary month total = Proservices + Procurement. Stop and report if it does not tie.
- Log the run in Data Sources. No notes in Summary!B2. No hardcodes outside Overrides.
- Return the updated workbook plus a short run report: per live month this week vs last week, every delta by invoice, new unmapped invoices, re-dated payments, open decisions (70000859 until I close it).
```

First run only, add: `Reproduce the 19/10/2026 baseline from standing-exceptions.md
before processing the new extract.`
