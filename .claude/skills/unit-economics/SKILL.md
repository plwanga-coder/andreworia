---
name: unit-economics
description: Builds a unit-economics and cohort .xlsx with live formulas across separate assumptions, cohort, and output tabs plus checks, computing CAC, LTV, LTV/CAC, and CAC payback from a monthly cohort retention grid. Use when someone needs unit economics, an LTV/CAC analysis, a payback model, or a cohort retention build. Do not trigger for editing an existing workbook through the Claude for Excel add-in.
---

# Unit Economics (LTV/CAC and Cohorts)

## When to use
Use when you need to know whether a customer is worth more than it costs to acquire: CAC, gross-margin contribution, retention, LTV, the LTV/CAC ratio, and CAC payback in months. Good for SaaS and subscription businesses, board reviews, and diligence where a cohort retention grid backs the headline ratios. The skill builds a downloadable .xlsx in the Claude app with openpyxl; it does not use any Excel add-in or Microsoft 365 connection.

## What it builds
A workbook with six tabs:
- Cover: title, color legend, scenario selector, headline LTV/CAC and payback.
- Assumptions: blue inputs (S&M spend, new customers, ARPU, gross margin %, churn, discount rate, scenario multipliers).
- Cohorts: a retention grid, cohorts down the rows and months across the columns, with retained % and revenue per cohort.
- Unit Economics: CAC, ARPU, contribution per customer, payback.
- LTV-CAC: LTV by the simple formula and by discounted cohort contribution, plus the ratio.
- Checks: consistency, monotonic retention, and payback sanity flags.

## Build workflow
1. Create the workbook and the six tabs in the order above.
2. On Assumptions, lay out all inputs in blue with units; add a scenario cell (1=Base, 2=Bull, 3=Bear) named `scn`.
3. Build CAC and per-customer contribution on Unit Economics from the spend and customer inputs.
4. Build the cohort grid: a month-0 base of 100% (or starting customers) decaying by retention across months.
5. Add a revenue and contribution row beneath the retention grid per cohort.
6. Build LTV two ways on LTV-CAC: closed-form and discounted cohort sum; compute the ratio and payback.
7. Build the Checks tab with consistency, monotonicity, and payback tests.
8. Recalculate the workbook headless (LibreOffice) so all formulas compute.
9. Verify zero formula errors (#REF!, #DIV/0!, #VALUE!, #NAME?); fix and re-recalculate in a loop, then deliver.

## Tab-by-tab spec

### Cover
- Title cell, model purpose line, and a build date.
- Color legend block: a blue swatch labelled "Input", black labelled "Formula", green labelled "Cross-tab link".
- Scenario echo: `=CHOOSE(scn,"Base","Bull","Bear")` pulling the active scenario from Assumptions.
- Headline outputs as green cross-tab links: LTV/CAC `=LTV-CAC!<ratio cell>`, CAC payback in months, and discounted LTV.
- A one-line read-me noting the workbook recalculates live in Excel and was verified error-free at build.

### Assumptions
- B2 Scenario selector (blue, 1/2/3), named `scn`.
- S&M spend per period B5; new customers acquired B6.
- ARPU per month B7; gross margin % B8.
- Monthly churn % B9 (so monthly retention `=1-B9`); discount rate per month B10.
- Scenario multipliers (Base/Bull/Bear) for churn and ARPU: active churn multiplier `=CHOOSE(scn,...)` in B12, active ARPU multiplier `=CHOOSE(scn,...)` in B13. Effective monthly churn is `B9*B12`.
- Consistency tolerance `tol` in B14 (for example 0.1%).

### Unit Economics
- CAC B3 `=Assumptions!B5/Assumptions!B6`.
- Effective ARPU B4 `=Assumptions!B7*Assumptions!B13`.
- Gross profit per customer-month B5 `=B4*Assumptions!B8`.
- Monthly retention B6 `=1-Assumptions!B9*Assumptions!B12`.
- CAC payback (months) B7 `=CAC/gross profit per customer-month` `=B3/B5`. If you prefer cohort-aware payback, compute cumulative gross profit per cohort month until it crosses CAC and report that month index.
- Average customer lifetime (months) B8 `=IF(Assumptions!B9*Assumptions!B12>0,1/(Assumptions!B9*Assumptions!B12),"n/a: zero churn")`.

### Cohorts
- Row 3 month headers 0,1,2,...,N across columns.
- Column A cohort labels (Cohort 1, Cohort 2, ...).
- Retained %: month 0 cell = 1 (100%); each later month `=prior_month_cell*'Unit Economics'!$B$6` (apply monthly retention). One consistent formula across the row.
- Active customers per cohort: `=retained% * Assumptions!$B$6`.
- Revenue per cohort-month: `=active customers * 'Unit Economics'!$B$4`.
- Contribution per cohort-month: `=revenue * Assumptions!$B$8`.

### LTV-CAC
- LTV simple (undiscounted, infinite horizon) B3: `=gross profit per customer-month / effective monthly churn`, in cells `='Unit Economics'!B5/(Assumptions!B9*Assumptions!B12)`. This is the familiar headline figure; guard churn = 0 with IF.
- LTV discounted-cohort B4: sum over months of contribution per customer discounted, `=SUMPRODUCT(contribution_per_customer_row, 1/(1+Assumptions!B10)^month_index)` using the month-0 cohort row divided by starting customers so it is per-customer.
- LTV/CAC B5 `=B4/'Unit Economics'!B3` (use the discounted LTV as primary).
- Payback echo B6 `='Unit Economics'!B7`.
- LTV discounted closed-form B7 (the same discounted, finite-horizon sum in one formula, used only to check the cohort grid): with q = retention/(1+discount) and M = number of month columns, `='Unit Economics'!B5*(1-q^M)/(1-q)`, in cells `='Unit Economics'!B5*(1-('Unit Economics'!B6/(1+Assumptions!B10))^COUNT(Cohorts!<month header row>))/(1-'Unit Economics'!B6/(1+Assumptions!B10))`.

### Checks
- Consistency: the cohort grid's discounted LTV (B4) equals the discounted closed-form (B7): `=IF(ABS(B4-B7)/B7<=tol,"PASS","FAIL")`. Do NOT compare against the simple LTV (B3): it is undiscounted and runs forever, so it is always higher by design (in the example, 1,000 vs about 678) and the check would always fail.
- Ordering sanity: `=IF('LTV-CAC'!B4<='LTV-CAC'!B3,"PASS","FAIL")` (discounted, finite-horizon LTV can never exceed the simple LTV).
- Monotonic retention within each cohort: `=IF(SUMPRODUCT(--(this_month_cells>prior_month_cells))=0,"PASS","FAIL")` so no month exceeds the prior.
- Payback positive and within a stated cap (blue input).
- LTV/CAC computed from the same ARPU, margin, and churn used in the cohort grid.

## Formula and formatting conventions
- Blue font for inputs (Assumptions only). Black for in-tab formulas. Green for cross-tab links.
- No hardcoded numbers in formulas; constants live on Assumptions. Use absolute refs to the retention and ARPU drivers so cohort rows fill cleanly.
- One consistent formula per cohort row so it copies across months without edits.
- Retention and margins as percent; CAC, LTV, ARPU as currency; payback and lifetime as months with one decimal.
- Name `scn`, `tol`, CAC, and the retention cell for readable Checks and LTV formulas.

## Checks
- LTV/CAC computed consistently (cohort-grid discounted LTV equals the discounted closed-form within tolerance; simple LTV is at least the discounted LTV).
- CAC payback reported in months and positive.
- Retention monotonically non-increasing within each cohort (no month rises above the prior).
- Margins and retention between 0 and 1; CAC and LTV non-negative.

## Recalculate and verify
After writing, recalculate headless with LibreOffice so openpyxl formula text becomes computed values. Scan every sheet for #REF!, #DIV/0!, #VALUE!, #NAME?. Note that LTV and payback divide by churn, so guard against churn = 0 producing #DIV/0!. Fix any offending formula or missing named range and recalculate again in a loop. Deliver only when the Checks tab shows all PASS and no error strings remain.

## Inputs to gather
- Sales and marketing spend and new customers in the same period (for CAC).
- ARPU per month and gross margin %.
- Monthly churn (or retention) and a discount rate.
- Number of cohorts and months to model in the grid.
- Scenario multipliers for churn and ARPU (Base/Bull/Bear) and the tolerance band.

## Example
Hypothetical: S&M 200,000, 400 new customers, so CAC 500. ARPU 50 per month, 80% gross margin gives 40 gross profit per customer-month, payback 12.5 months. Churn 4% monthly implies 25-month average lifetime; simple LTV 1,000 (LTV/CAC 2.0x on that basis). With a 1% monthly discount rate over a 36-month grid, the discounted cohort LTV is about 678 (LTV/CAC about 1.36x), and the discounted closed-form matches it exactly, so the consistency check passes. The cohort grid decays each cohort 100%, 96%, 92.2%, ... across months, and contribution sums to the discounted LTV.
