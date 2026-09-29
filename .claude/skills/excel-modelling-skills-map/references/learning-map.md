# Learning map: how to obtain the 12 skills

## 1. Stage ladder

| Stage | Name | Competencies | Hours | Exit test |
|:-:|:--|:--|:-:|:--|
| 0 | Excel fluency | references, lookups, SUMIFS, logic, keyboard, Tables | 10-20 | Diagnostic Stage 0 in SKILL.md |
| 1 | Data hygiene | data-cleaning-for-excel, pivot-table-builder | 10-16 | Clean a messy export plus log; pivot total reconciles |
| 2 | Structure | model-architecture-template, inputs-calcs-outputs-design | 12-18 | Restructured model, zero hardcodes in Calcs, Checks PASS |
| 3 | Driver builds | revenue-build, unit-economics | 20-30 | Roll-forward ties every period; LTV/CAC explained |
| 4 | Stress testing | scenario-manager, sensitivity-tables, sensitivity-tornado | 17-24 | Live selector; native 2-way Data Table; ranked tornado |
| 5 | Assurance and handover | formula-audit-checker, assumption-registry-builder, output-summary-tab | 26-43 | Audit log on a model you did not build; register; one-page summary |
| C | Capstone | all 12 in one model | 15-25 | Passes an independent audit with zero Critical/Major findings |

## 2. Dependency graph

```
Stage 0  Excel fluency
            |
Stage 1  data-cleaning-for-excel ----> pivot-table-builder
            |
Stage 2  model-architecture-template <-> inputs-calcs-outputs-design
            |                                  |
Stage 3  revenue-build --------------> unit-economics
            |                                  |
Stage 4  scenario-manager --> sensitivity-tables --> sensitivity-tornado
            |                                            |
Stage 5  formula-audit-checker    assumption-registry-builder (needs tornado impacts)
                     \                    /
                      output-summary-tab
                             |
                         Capstone
```

Hard prerequisites: Stage 2 before any build; `sensitivity-tables` before
`sensitivity-tornado`; `sensitivity-tornado` before `assumption-registry-builder`.
`formula-audit-checker` can start after Stage 2 as a habit, but mastery needs Stage 4.

## 3. Competency cards

Each card lists the Excel features to learn, a drill, pass criteria (what the coach
checks) and resources. Standard method for every drill: **build it by hand first, then
run the matching skill on the same inputs, then diff the two.**

### Stage 0: Excel fluency (prerequisite, not in the pack)
- **Learn:** relative/absolute/mixed references (F4); XLOOKUP, INDEX/MATCH; SUMIFS,
  COUNTIFS, AVERAGEIFS; IF/IFS/IFNA/AND/OR; EDATE, EOMONTH; Excel Tables; named ranges;
  keyboard navigation; custom number formats.
- **Drill:** build a 12-month x 5-product sales summary from a transaction list using only
  SUMIFS keyed on month start and end dates. Recreate it with a mouse-free workflow.
- **Pass:** every formula fills across and down without edits; totals reconcile; the drill
  is done in under 30 minutes.
- **Resources:** Microsoft Learn / Microsoft Support Excel training; ExcelJet function
  guides; Microsoft Office Specialist Excel Associate (MO-210) syllabus as a checklist.

### Stage 1a: data-cleaning-for-excel
- **Learn:** TRIM + SUBSTITUTE(CHAR(160)), NUMBERVALUE, DATEVALUE, Text to Columns,
  Remove Duplicates, Flash Fill; **Power Query** (Get & Transform: change type with
  locale, replace values, remove rows, merge queries).
- **Drill:** take a real export (bank, CRM, mobile-money or ERP) with at least three
  problems: mixed dates, text numbers, duplicate rows or subtotal rows. Produce Raw /
  Clean / Change Log tabs. Then redo the clean in Power Query so it refreshes.
- **Pass:** Raw untouched; rows in = rows out + rows removed (logged); no text-numbers
  left (`=SUMPRODUCT(--ISTEXT(range))=0`); IDs keep leading zeros; ambiguous date order
  resolved by asking, not guessing; the Power Query version refreshes on new data.
- **Resources:** Microsoft Learn Power Query documentation; *M is for (Data) Monkey*
  (Puls & Escobar); Microsoft Office Specialist Excel Expert (MO-211) data-prep objectives.

### Stage 1b: pivot-table-builder
- **Learn:** Ctrl+T Tables, PivotTables, value field settings, date grouping, slicers,
  GETPIVOTDATA; the equivalent SUMIFS grid.
- **Drill:** from your Stage 1a clean table, answer one business question ("which region
  grew fastest?") with a native pivot, then with a SUMIFS grid.
- **Pass:** layout justified by the question; grand total equals the source total (shown
  in a check cell); both versions agree to the cent.
- **Resources:** Microsoft Support "Create a PivotTable"; ExcelJet pivot table guides.

### Stage 2: model-architecture-template + inputs-calcs-outputs-design (study together)
- **Learn:** tab design with a flow map; colour convention (choose blue font *or* blue
  fill for inputs and document it); cell styles; custom formats (`#,##0;(#,##0);-`,
  `0.0%`, `0.0x`); named ranges; sheet protection; print titles; a Checks tab with
  `COUNTIF(range,"FAIL*")`.
- **Drill:** take a workbook you already use (budget, month-end, a forecast) and rebuild it
  as Cover / Legend / Inputs / Calc tabs / Checks / Outputs / Documentation.
- **Pass:** zero numeric constants in Calc formulas (proved with a FORMULATEXT or
  script scan, not by eye); every input labelled with a unit; the Outputs tab contains
  only references; the Checks traffic light reads ALL CHECKS PASS and turns red when you
  deliberately break a tie-out; version table present.
- **Resources:** **FAST Standard** (fast-standard.org, free); **ICAEW Financial Modelling
  Code** (free); SMART / Operis modelling guidelines; *Financial Modelling in Practice*
  (Michael Rees).

### Stage 3a: revenue-build
- **Learn:** period timeline row, roll-forward (corkscrew) schedules, one-formula-per-row,
  CHOOSE for scenario multipliers, annualisation (SUMIFS by year), checks for roll-forward
  identity and continuity.
- **Drill:** 36-month customer build: starting customers, adds = spend / CAC, monthly
  churn, ARPU. Add a units x price variant for a two-product business relevant to you.
- **Pass:** every period ties (ending = beginning + adds - churn), shown on Checks; the row
  formula is identical across all columns (`Ctrl+\` row differences selects nothing); your
  month-1 figures match `revenue-build`'s reference (1,000 + 200 - 30 = 1,170 on its
  example inputs).
- **Resources:** *Financial Modeling* (Simon Benninga) for technique; FP&A-oriented
  courses (e.g. CFI FMVA modules on forecasting).

### Stage 3b: unit-economics
- **Learn:** CAC, contribution margin, CAC payback, simple LTV = GP per month / churn,
  discounted LTV, cohort (triangle) grids, SUMPRODUCT with discount factors.
- **Drill:** build a 12-cohort x 36-month retention grid from inputs, compute both LTVs and
  LTV/CAC, then load real cohort data if you have it.
- **Pass:** discounted cohort LTV equals the discounted closed form within tolerance;
  simple LTV >= discounted LTV; you can explain in two sentences why LTV/CAC falls from
  2.0x to ~1.36x on the skill's example; retention never rises within a cohort.
- **Resources:** David Skok, "SaaS Metrics 2.0" (forEntrepreneurs.com); a16z / Bessemer
  published SaaS metric primers.

### Stage 4a: scenario-manager
- **Learn:** data validation lists, CHOOSE vs INDEX, the flow selector -> active column ->
  model, parallel-calc comparison, directional sanity checks. Also learn why Excel's
  built-in Scenario Manager is avoided.
- **Drill:** add a Base/Bull/Bear layer to your Stage 3a model for at least four drivers,
  with a side-by-side output table.
- **Pass:** one selector cell controls everything; no scenario-driven assumption is typed
  into the model; Bull >= Base >= Bear on the headline metric; toggling 1/2/3 changes
  outputs and all checks stay PASS.
- **Resources:** FAST Standard (scenario sections); ExcelJet CHOOSE/INDEX guides.

### Stage 4b: sensitivity-tables
- **Learn:** Data > What-If Analysis > Data Table; row vs column input cell; same-sheet
  rule; odd step counts centred on base; calculation mode "Automatic except for data
  tables"; conditional formatting colour scales.
- **Drill:** on the model sheet, build a one-way table (discount rate) and a two-way table
  (growth x discount rate) on NPV or ending ARR. Then deliberately swap the input cells and
  observe the silent transposition.
- **Pass:** the centre cell equals the live output; axes labelled with the driver and cell
  address; you can explain the swap error. Note: this must be done in Excel. Claude's
  version is a computed snapshot and Google Sheets has no Data Tables.
- **Resources:** Microsoft Support "Calculate multiple results by using a data table";
  Wall Street Prep / Macabacus free sensitivity-table guides.

### Stage 4c: sensitivity-tornado
- **Learn:** one-at-a-time swings, choosing credible ranges, ABS impact ranking, clustered
  or stacked bar charts with reversed category order, a zero reference line.
- **Drill:** tornado on 8-10 drivers of your model's key output, with a swing-rationale
  column.
- **Pass:** sorted by absolute range; the most important driver sits at the top of the
  chart; each range has a written rationale; you can name the top three drivers and what
  you would do to firm them up.
- **Resources:** Peltier Technical Services (tornado chart tutorials).

### Stage 5a: formula-audit-checker
- **Learn:** Error Checking, Evaluate Formula (step through), Trace Precedents and
  Dependents, Watch Window, Inquire add-in (Workbook Analysis, Compare Files), FORMULATEXT
  scans, IFNA vs IFERROR, Name Manager hygiene, sign conventions, severity grading.
- **Drill:** audit a model you did not build (a colleague's, a public template, or one
  with seeded errors: ask Claude to plant 10 errors in a copy and not tell you where).
  Produce a findings log.
- **Pass:** at least 8 of 10 seeded errors found; each finding has tab, cell, type,
  severity, fix; overall status stated using the skill's Pass / Conditional / Fail rule.
- **Resources:** **EuSpRIG** (European Spreadsheet Risks Interest Group) papers and horror
  stories; ICAEW Code review sections; *Spreadsheet Check and Control* (Patrick O'Beirne).

### Stage 5b: assumption-registry-builder
- **Learn:** RANK on ABS impact, COUNTIF/COUNTBLANK summaries, conditional formatting with
  `TODAY()-date>90`, AutoFilter, sign-off columns.
- **Drill:** build the register for your Stage 4 model, sourcing every material assumption
  and linking Base/Bull/Bear to the scenario inputs and impact to the tornado.
- **Pass:** no source reads "assumed" or is blank (`COUNTBLANK` = 0); base values are
  references, not retyped; rank 1 matches the tornado's top driver; every row has an owner.
- **Resources:** ICAEW Code (documentation principles); FAST Standard.

### Stage 5c: output-summary-tab
- **Learn:** page layout, print areas, fit to one page, KPI tile formatting, gridlines off,
  hyperlinks; **Pyramid Principle** writing (answer first, then support).
- **Drill:** one-page summary for your Stage 4 model: 4-6 KPI tiles, a 100-200 word
  narrative, a Bear/Base/Bull table, and an assumption log.
- **Pass:** only references on the tab (no calculations); prints on one landscape page;
  the headline sentence states the answer and a number; a reader who has not seen the
  model can say the conclusion from a screenshot alone.
- **Resources:** *The Pyramid Principle* (Barbara Minto); *Storytelling with Data*
  (Cole Nussbaumer Knaflic).

### Capstone
Build one model for a business you know (your employer, a client, a family business, a
public company) using all 12 competencies: clean source data -> architecture -> drivers
-> unit economics if relevant -> scenarios -> sensitivities -> tornado -> self-audit ->
register -> one-page summary. Then ask someone else (or Claude with
`formula-audit-checker`) to audit it. **Pass: zero Critical or Major findings.** This model
is your portfolio piece for interviews.

## 4. Schedules

**Intensive (~12 weeks at 10-12 h/week)**
| Weeks | Focus |
|:--|:--|
| 1 | Stage 0 gaps + Stage 1 |
| 2-3 | Stage 2 |
| 4-6 | Stage 3 |
| 7-9 | Stage 4 |
| 10-11 | Stage 5 |
| 12 | Capstone |

**Part-time (~26 weeks at 5 h/week)**: double each block above; start the capstone in
week 22. Keep auditing (Stage 5a habits) running from week 6 on every model you touch.

## 5. Role emphasis

| Target role | Prioritise | Can defer |
|:--|:--|:--|
| FP&A / corporate finance | revenue-build, scenario-manager, output-summary-tab, assumption register | unit-economics (unless subscription) |
| Investment banking / PE / corp dev | sensitivity-tables, formula audit, architecture; then add 3-statement, DCF, LBO (beyond the pack) | pivot, unit-economics |
| Startup operator / VC | unit-economics, revenue-build, scenario-manager | assumption register depth |
| Consulting / strategy | output-summary-tab, tornado, scenarios, data cleaning + pivots | deep audit tooling |
| Audit / model review | formula-audit-checker, architecture, register | building speed |
| SME owner / personal finance | data cleaning, pivots, inputs-calcs-outputs, scenarios | Data Tables, tornado |

## 6. Beyond the pack (next skills after the capstone)
- **Three-statement integration**: P&L, balance sheet and cash flow linked, with a balance
  check. (If available, the `excel-model-builder` skill builds one to compare against.)
- **Valuation**: DCF (WACC, terminal value), trading and transaction comparables; LBO with a
  debt schedule and returns waterfall.
- **Operational schedules**: working capital days, fixed assets and depreciation, tax,
  timing flags, and a 13-week direct cash flow.
- **Modern Excel**: dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE), LET, LAMBDA; Power
  Pivot and DAX for large data.
- **Analysis**: variance bridges (price/volume/mix), rolling forecasts, Monte Carlo
  simulation.
- **Automation**: Python in Excel or openpyxl, which is how the four build skills in this
  pack generate workbooks.

## 7. Credentials (optional, to signal the skills)
| Credential | Body | Best for | Covers |
|:--|:--|:--|:--|
| Microsoft Office Specialist: Excel Associate (MO-210) / Expert (MO-211) | Microsoft | Proving Excel fluency (Stage 0-1) | Functions, pivots, data tools, what-if analysis |
| Financial Modeling & Valuation Analyst (FMVA) | Corporate Finance Institute | Structured course path | Stages 2-4 plus 3-statement and valuation |
| Advanced Financial Modeler (AFM), then Chartered Financial Modeler (CFM) | Financial Modeling Institute | Practical, exam-based proof of modelling | Build a model under time pressure, graded on accuracy and structure |
| FAST Standard certification | FAST Standard Organisation | Standards-driven modelling roles | Structure and conventions (Stage 2, 5) |

Do the capstone before any practical exam: AFM-style tests reward exactly the
architecture, checks and speed that the capstone builds.

## 8. How to obtain the Claude skills themselves
To use the 12 skills as sparring partners:
- **Claude Code:** they load automatically in any session opened on this repo
  (`.claude/skills/`), or copy them with
  `cp -r .claude/skills/* ~/.claude/skills/`.
- **Claude app:** Settings -> Capabilities -> Skills, and upload each skill folder (or a zip).
- **Build skills need:** `pip install openpyxl numpy` and `libreoffice-calc` for the
  recalculation check.
