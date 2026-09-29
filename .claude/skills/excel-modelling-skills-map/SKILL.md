---
name: excel-modelling-skills-map
description: Coaches a person to acquire the 12 Excel financial-modelling competencies in this pack (data cleaning, pivots, model architecture, inputs-calcs-outputs separation, revenue builds, unit economics, scenarios, sensitivity tables, tornado analysis, formula audit, assumption registers, executive summary tabs). Diagnoses their current level, places them on a staged learning map, prescribes drills with pass criteria, and points to recognised standards and certifications. Use when someone asks how to learn or get better at financial modelling in Excel, wants a training plan or skills roadmap, asks which modelling skill to learn next, or wants their modelling skills assessed.
---

# Excel Modelling Skills Map

## When to use
Use this when a person (not Claude) wants to *obtain* the competencies that the other
12 skills in this pack perform: "how do I learn to build models like this?", "give me a
training plan", "what should I learn next?", "assess my modelling", "I'm a junior analyst,
where do I start?". Do not use it to build a model; hand that to the relevant build skill.

## Reference files
- `references/appraisal.md`: what each of the 12 skills is, what competency it really
  tests, its prerequisites, difficulty, career value, and how good the skill file is as a
  teacher (with known limitations). Read it before recommending a skill as a study aid.
- `references/learning-map.md`: the six-stage roadmap, the dependency graph, and for every
  competency the Excel features to master, a drill, pass criteria, time estimate and
  external resources. Also covers gaps the pack does not teach, and certification routes.

## Method
1. **Establish the goal and constraints.** Ask (if not stated): role or target role (FP&A,
   investment banking/PE, consulting, startup operator, SME owner, audit), Excel version
   (Microsoft 365, 2019, Google Sheets, LibreOffice), hours per week available, and a
   deadline if any. The goal decides which branch of the map matters most (see
   "Role emphasis" in `learning-map.md`).
2. **Diagnose the current level.** Run the diagnostic below. Ask the person to answer
   honestly or, better, to do the task. Place them at the first stage where they cannot
   pass every item. Do not skip a stage because the person claims seniority. Stage 0 gaps
   break everything downstream.
3. **Show the map.** Present the stage ladder and the dependency graph from
   `learning-map.md`, marking where they are and the next two competencies.
4. **Prescribe the next block.** For each of the next 1-3 competencies give: why it
   matters (from `appraisal.md`), the Excel features to learn, the drill, the pass
   criteria, the time estimate, and one or two resources. Do not prescribe more than three
   at once.
5. **Use the pack as a sparring partner, not a crutch.** For each drill: the person builds
   it first, by hand. Only then run the matching skill (e.g. `revenue-build`) on the same
   inputs to produce a reference answer, and compare the two cell by cell. Differences
   are the lesson. Where `appraisal.md` records a limitation (Data Tables, native pivots,
   LibreOffice), say so up front.
6. **Assess against the pass criteria.** When the person shares their workbook or
   answers, grade each criterion PASS/FAIL with the specific cell or reason. Run
   `formula-audit-checker` on their file as the independent check from Stage 3 onward.
   Advance them only when every criterion passes.
7. **Keep a progress log.** End each session with a short table: competency | status
   (Not started / In progress / Passed) | evidence (file or date) | next action. Ask the
   person to paste it back next time so you can resume.

## Diagnostic (place the person on the map)
Stage 0 -- Excel fluency
- Can explain and use `$A$1`, `A$1`, `$A1` without trial and error.
- Writes `XLOOKUP` or `INDEX/MATCH`, `SUMIFS`, `IF`/`IFS`, `IFNA` from memory.
- Navigates by keyboard (Ctrl+arrows, Ctrl+Shift+arrows, F2, F4, Ctrl+PgUp/PgDn).

Stage 1 -- Data hygiene
- Given a messy export, can produce a clean table plus a log of what changed.
- Can build a PivotTable and prove its grand total equals the source total.

Stage 2 -- Structure
- Can lay out Cover / Inputs / Calcs / Checks / Outputs and state the colour convention.
- No hardcoded numbers in calculation formulas in their last model.

Stage 3 -- Driver-based builds
- Can build a beginning + adds - churn = ending roll-forward that ties every period.
- Can compute CAC, LTV (simple and discounted) and payback, and say why they differ.

Stage 4 -- Stress testing
- Can build a live Base/Bull/Bear selector with `CHOOSE`/`INDEX`.
- Can build a two-variable Data Table and name which cell is the row vs column input.
- Can rank drivers with a tornado and explain the top three.

Stage 5 -- Assurance and handover
- Can audit someone else's model and produce a severity-graded findings log.
- Can produce an assumption register with sources and owners, and a one-page summary
  tab that passes the "screenshot test".

## Output format
1. Goal and constraints (one line each).
2. Diagnostic result: current stage, with the specific items passed and failed.
3. Map excerpt: the stage ladder with "you are here" and the next two competencies.
4. Next block (1-3 competencies): why, features, drill, pass criteria, hours, resources.
5. Progress log table.

## Example
**User says:** "I'm an accountant, comfortable with VLOOKUP and pivots. I want to move
into FP&A in six months. About five hours a week."

**Claude responds (abridged):**
- Diagnostic: Stage 0 passed except `XLOOKUP` and mixed references (fix first, ~4 h).
  Stage 1 passed (pivots, cleaning); verify pivot total check. Placed at **Stage 2**.
- Next block: `model-architecture-template` + `inputs-calcs-outputs-design` (together,
  ~10 h). Drill: restructure one of your own month-end workbooks into Inputs / Calcs /
  Checks / Outputs. Pass: zero hardcodes in Calcs (proved by a FORMULATEXT scan), every
  input blue and labelled with a unit, Checks tab shows ALL CHECKS PASS.
- Then `revenue-build` (~12 h), prioritised because FP&A lives on driver-based forecasts.
- Six-month route: Stages 2-4 in weeks 1-16, Stage 5 in weeks 17-22, portfolio model in
  weeks 23-26. Optional credential: FMI Advanced Financial Modeler after the portfolio.
