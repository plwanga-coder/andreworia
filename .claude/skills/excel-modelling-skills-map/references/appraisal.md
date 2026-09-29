# Appraisal of the 12 skills

Source: the 12 `SKILL.md` files in `.claude/skills/` and `SKILLS_AUDIT.md` (security audit
of 25-26 Sep 2026: SAFE; 11 technical findings, all patched).

## 1. Identification

| # | Skill | Type | Cluster | Competency it represents |
|--:|:--|:--|:--|:--|
| 1 | data-cleaning-for-excel | Method | Data hygiene | Turning a raw export into analysable data, with an audit trail |
| 2 | pivot-table-builder | Method | Data hygiene | Summarising flat data by segment and period, and proving the total |
| 3 | model-architecture-template | Method | Structure | Designing the workbook before building: tabs, colours, formats, checks, versioning |
| 4 | inputs-calcs-outputs-design | Method | Structure | Separating assumptions, logic and presentation; the no-hardcode rule |
| 5 | revenue-build | Builds .xlsx | Driver builds | Bottom-up forecasting from operating drivers (roll-forward schedules) |
| 6 | unit-economics | Builds .xlsx | Driver builds | CAC, LTV, LTV/CAC, payback and cohort retention analysis |
| 7 | scenario-manager | Builds .xlsx | Stress testing | Visible, auditable Base/Bull/Bear switching with CHOOSE/INDEX |
| 8 | sensitivity-tables | Builds .xlsx | Stress testing | One- and two-variable sensitivity grids and Excel Data Tables |
| 9 | sensitivity-tornado | Method | Stress testing | Ranking which assumptions matter; tornado chart construction |
| 10 | formula-audit-checker | Method | Assurance | Reviewing a model for errors, hardcodes, circularity and masked errors |
| 11 | assumption-registry-builder | Method | Assurance | Documenting provenance, ranges, owners and sign-off for every assumption |
| 12 | output-summary-tab | Method | Handover | Communicating the result on one page: KPI tiles, narrative, sensitivities |

The 12 form one pipeline (the README's recommended chain):
clean -> structure -> build drivers -> stress -> audit -> document -> present.

## 2. Analysis and appraisal per skill

Scales: **Difficulty** 1 (a day of practice) to 5 (months). **Career value** 1 (nice to
have) to 5 (non-negotiable for a modelling role). **Teaching quality** is how well the
`SKILL.md` works as study material, after the patches.

### 1. data-cleaning-for-excel
- **What it really tests:** discipline more than technique. Never clean in place; profile
  before changing; distinguish blank from zero; keep IDs as text; ask when day/month order
  is ambiguous; log every change.
- **Prerequisites:** TRIM, SUBSTITUTE, VALUE/NUMBERVALUE, DATEVALUE, Remove Duplicates;
  ideally Power Query.
- **Difficulty 2 | Career value 5 | Time to competence ~6-10 h.**
- **Teaching quality: Good (was thin, expanded in the patch).** The rule list and worked
  example are sound. It is weak on Power Query, the tool a professional would actually use
  for repeatable cleans, which is mentioned in one line.
- **Verdict:** foundational and underrated. Most model errors start in the data. Learn first.

### 2. pivot-table-builder
- **What it tests:** choosing the layout from the question being asked, spotting fields
  that will not aggregate, and reconciling the grand total to the source.
- **Prerequisites:** Excel Tables (Ctrl+T), PivotTables, date grouping, SUMIFS.
- **Difficulty 1-2 | Career value 4 | Time ~4-6 h.**
- **Teaching quality: Adequate.** Short. The SUMIFS-grid alternative is a valuable lesson
  in itself (live formula summaries are more auditable than pivots).
- **Limitation:** when Claude builds the file, openpyxl cannot create a native pivot. It
  produces a SUMIFS grid or a static pandas summary instead. Learners should still build
  native pivots by hand.
- **Verdict:** quick win; pair with data cleaning.

### 3. model-architecture-template
- **What it tests:** planning. Tab inventory with "fed by / feeds into", colour code,
  number formats, Checks tab with a master traffic light, version table, print setup.
- **Prerequisites:** cell styles, custom number formats, named ranges, hyperlinks,
  COUNTIF with wildcards.
- **Difficulty 2 | Career value 5 | Time ~6-8 h** (the habit takes several models).
- **Teaching quality: Very good.** Detailed and opinionated in line with industry
  practice. The COUNTIF `"FAIL*"` wildcard bug is patched. The colour scheme differs
  slightly from the build skills (blue *fill* here, blue *font* in the build skills), so a
  learner should pick one convention and document it.
- **Verdict:** highest leverage structural skill. Teaches "design before you build".

### 4. inputs-calcs-outputs-design
- **What it tests:** the single most important modelling discipline: separating inputs,
  calculations and outputs, with no hardcoded numbers in calculations.
- **Prerequisites:** cross-sheet references, absolute references, named ranges, sheet
  protection.
- **Difficulty 2 | Career value 5 | Time ~6-10 h.**
- **Teaching quality: Very good.** Clear rules, a conventions checklist and a worked ARR
  example. It overlaps heavily with #3; study them together.
- **Verdict:** non-negotiable. Everything from Stage 3 onward assumes it.

### 5. revenue-build
- **What it tests:** driver-based forecasting. A roll-forward (beginning + adds - churn =
  ending), revenue = units x price, one consistent formula per row, a scenario multiplier,
  and roll-forward integrity checks.
- **Prerequisites:** #3, #4; SUM over periods, CHOOSE, absolute references.
- **Difficulty 3 | Career value 5 | Time ~10-15 h.**
- **Teaching quality: Very good.** Worked example verified (1,000 + 200 - 30 = 1,170).
- **Limitations:** the sensitivity grid is computed in Python as a snapshot (a native Data
  Table cannot be written from code). Recalculation assumes LibreOffice.
- **Verdict:** the core FP&A skill. The roll-forward pattern also transfers to headcount,
  inventory, debt and fixed assets.

### 6. unit-economics
- **What it tests:** CAC, contribution margin, CAC payback, simple and discounted LTV,
  cohort retention grids, and the difference between them.
- **Prerequisites:** #5; SUMPRODUCT, discounting, geometric series.
- **Difficulty 3-4 | Career value 4** (5 in SaaS/subscription or VC) **| Time ~10-15 h.**
- **Teaching quality: Very good after patch.** The consistency check now compares like
  with like (discounted cohort LTV vs discounted closed form: 678.10 = 678.10). The lesson
  that simple LTV (1,000) overstates discounted LTV (~678), dropping LTV/CAC from 2.0x to
  ~1.36x, is worth learning on its own.
- **Verdict:** high value for startup, SaaS and investor roles; optional for corporate FP&A.

### 7. scenario-manager
- **What it tests:** one selector cell driving every scenario-dependent assumption via
  CHOOSE/INDEX; a parallel-calc comparison that shows all cases at once; directional
  sanity (Bull must beat Base).
- **Prerequisites:** #4; CHOOSE, INDEX, data validation lists.
- **Difficulty 2-3 | Career value 5 | Time ~6-8 h.**
- **Teaching quality: Very good.** It explains clearly why Excel's built-in Scenario
  Manager is inferior (hidden, overwrites inputs).
- **Verdict:** essential. Works in Excel and Google Sheets.

### 8. sensitivity-tables
- **What it tests:** one- and two-variable grids centred on the base case with odd step
  counts, correct row vs column input cells, and checking the base-case cell.
- **Prerequisites:** #4, #7; Data > What-If Analysis > Data Table; conditional formatting.
- **Difficulty 3 | Career value 5** (banking, PE, corporate development) **| Time ~6-8 h.**
- **Teaching quality: Good, with an important caveat.** The original approach was broken
  (high-severity finding): native Data Tables need their input cells on the same sheet as
  the table, and openpyxl cannot write them. The patched skill gives the grid as a
  computed snapshot and tells the user how to add a live table in Excel. Learners must
  practise the native Data Table by hand; that is the interview skill.
- **Limitation:** Data Tables do not exist in Google Sheets.
- **Verdict:** essential for valuation work; learn by hand, not via Claude.

### 9. sensitivity-tornado
- **What it tests:** one-at-a-time sensitivity, defining credible swing ranges, ranking
  by absolute impact, and building a tornado chart from a stacked or clustered bar chart.
- **Prerequisites:** #8; bar-chart formatting, sorting, ABS.
- **Difficulty 3 | Career value 4 | Time ~5-8 h.**
- **Teaching quality: Good.** The example table order is fixed. The chart-building advice
  is correct but brief. Expect to experiment.
- **Verdict:** the best tool for answering "which assumption matters?". Feeds #11.

### 10. formula-audit-checker
- **What it tests:** reviewing someone else's model. Error checking, circular references,
  trace precedents/dependents, hunting hardcodes inside formulas (FORMULATEXT, Inquire,
  script), IFERROR masking vs IFNA, named ranges, sign conventions, severity grading.
- **Prerequisites:** #3, #4 and experience of real models.
- **Difficulty 4 | Career value 5 | Time ~15-25 h** (and years to master).
- **Teaching quality: Very good after patch.** Two inaccuracies are fixed: Go To Special
  cannot find constants inside formulas, and Ctrl+[ does not draw arrows. The findings log
  and Pass/Conditional/Fail rule are professional-grade.
- **Verdict:** separates juniors from seniors. It is also the tool this map uses to grade
  learners' work.

### 11. assumption-registry-builder
- **What it tests:** provenance and governance. Every material assumption gets a source,
  Bull/Bear range, sensitivity rank (RANK on ABS impact), owner, review date and sign-off.
- **Prerequisites:** #9 (the impact column comes from the tornado); RANK, COUNTIF,
  COUNTBLANK, conditional formatting with TODAY().
- **Difficulty 2-3 | Career value 4** (5 in corporate/board/regulated settings) **| Time ~5-8 h.**
- **Teaching quality: Very good after patch** (the RANK-on-signed-impact bug and the
  threshold inconsistency are fixed). The "acceptable vs unacceptable source" guidance is
  especially useful.
- **Verdict:** turns a model into an asset someone else can trust.

### 12. output-summary-tab
- **What it tests:** communication. A five-zone one-page layout, 4-6 KPI tiles, a
  Pyramid Principle narrative (answer first), compact sensitivities, and the
  "screenshot test".
- **Prerequisites:** #4; page layout, print areas, formatting; business writing.
- **Difficulty 2 (Excel) / 4 (writing) | Career value 5 | Time ~6-10 h.**
- **Teaching quality: Very good.**
- **Verdict:** the part stakeholders actually see. The writing matters as much as the Excel.

## 3. Summary scorecard

| Skill | Difficulty | Value | Hours | Teaching quality | Learn it via Claude? |
|:--|:-:|:-:|:-:|:--|:--|
| data-cleaning-for-excel | 2 | 5 | 6-10 | Good | Yes, then redo in Power Query |
| pivot-table-builder | 1-2 | 4 | 4-6 | Adequate | Build native pivots by hand |
| model-architecture-template | 2 | 5 | 6-8 | Very good | Yes |
| inputs-calcs-outputs-design | 2 | 5 | 6-10 | Very good | Yes |
| revenue-build | 3 | 5 | 10-15 | Very good | Build first, compare to Claude's |
| unit-economics | 3-4 | 4 | 10-15 | Very good | Build first, compare |
| scenario-manager | 2-3 | 5 | 6-8 | Very good | Build first, compare |
| sensitivity-tables | 3 | 5 | 6-8 | Good (caveat) | Native Data Table by hand only |
| sensitivity-tornado | 3 | 4 | 5-8 | Good | Yes, chart by hand |
| formula-audit-checker | 4 | 5 | 15-25 | Very good | Use as grader and checklist |
| assumption-registry-builder | 2-3 | 4 | 5-8 | Very good | Yes |
| output-summary-tab | 2 / 4 | 5 | 6-10 | Very good | Write the narrative yourself |

Total to working competence: roughly **90-140 hours**, plus Stage 0 fluency if missing
(10-20 h) and a capstone model (15-25 h).

## 4. Overall appraisal of the pack

**Strengths**
- It is a coherent pipeline, not a random list; the order matches how professional models
  are built and reviewed.
- It emphasises the disciplines that prevent real losses: no hardcodes, checks tabs,
  visible scenarios, audit logs, sourced assumptions. These match published standards (FAST
  Standard, ICAEW Financial Modelling Code).
- The worked examples are arithmetically correct (verified in the audit).
- It is safe: plain Markdown, no scripts, no network calls.

**Weaknesses**
- The build skills lean on Python/openpyxl and LibreOffice. That shows Claude how to build
  a file, but it hides the Excel mechanics (Data Tables, native pivots) that a learner
  must still practise by hand.
- Colour conventions differ slightly between the method skills (blue fill) and the build
  skills (blue font); a learner must choose one convention.
- The two data skills are short compared with the rest.
- Examples are SaaS-heavy. Retail, manufacturing, project finance and public-sector
  learners need to translate the drivers.

**Gaps (not covered by any of the 12; see learning-map.md, "Beyond the pack")**
- Three-statement integration (P&L, balance sheet, cash flow that balances).
- Valuation: DCF, WACC, terminal value, comparables; LBO/debt schedules.
- Working capital, fixed-asset and tax schedules; timing flags and period switches.
- Modern Excel: dynamic arrays, LET, LAMBDA, Power Query, Power Pivot/DAX.
- Charts and dashboards beyond KPI tiles; probabilistic methods (Monte Carlo).
- Version control, model change management and peer review workflow.
