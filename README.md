# Claude Excel Skills

**12 Claude skills for building financial models and clean workbooks that survive a review.**

These cover the discipline behind a model that someone else can open: a tab structure agreed before the first formula, inputs separated from calculations, a driver-based revenue build, a scenario toggle you can actually audit, and a formula check run before the model reaches an investment committee. Underneath it all is the unglamorous part — cleaning the pasted export and pivoting it — because that is where most models really start.

Several of these skills build a real `.xlsx` from scratch through Claude using `openpyxl`, with live formulas and a recalculate-and-verify pass (headless LibreOffice), and hand you the file. They say so explicitly, and they say equally explicitly that they are **not** the Claude for Excel add-in. The rest are method skills: structure, conventions and audit routines that make Claude reason about a workbook the way a modeller does.

The skills in this repo have been audited and patched. See [SKILLS_AUDIT.md](SKILLS_AUDIT.md) for the security review, the technical findings and what was changed.

## The skills

### Structure — lay out the workbook before a formula goes in

| Skill | What it does |
|:--|:--|
| [`model-architecture-template`](.claude/skills/model-architecture-template) | Sets the master tab set, naming conventions, colour codes, number formats and print setup for a new model — the empty shell every other skill fills. |
| [`inputs-calcs-outputs-design`](.claude/skills/inputs-calcs-outputs-design) | Enforces the three-zone separation: an Inputs tab of editable assumptions, Calculation tabs with no hardcodes, and an Outputs tab that only references. |
| [`assumption-registry-builder`](.claude/skills/assumption-registry-builder) | Builds the assumption register — name, category, base value, source, Bull and Bear cases, sensitivity rank and owner — so every number has a provenance and a signature. |

### Build — clean the data, then drive the numbers

| Skill | What it does |
|:--|:--|
| [`data-cleaning-for-excel`](.claude/skills/data-cleaning-for-excel) | Turns a pasted export with mixed date formats, text numbers, duplicates and stray whitespace into a consistent range, with a change log so the cleaning is auditable. |
| [`pivot-table-builder`](.claude/skills/pivot-table-builder) | Specifies and builds a pivot from a flat range — rows, columns, values, filters — and flags the fields that will not aggregate cleanly. |
| [`revenue-build`](.claude/skills/revenue-build) | Builds a driver-based revenue `.xlsx`: a customer or units schedule of beginning, adds, churn and ending rolled into revenue, across drivers, build and output tabs. |
| [`unit-economics`](.claude/skills/unit-economics) | Builds a unit-economics and cohort `.xlsx` — CAC, LTV, LTV/CAC and CAC payback computed off a monthly retention grid rather than asserted. |

### Stress it, audit it, hand it over

| Skill | What it does |
|:--|:--|
| [`scenario-manager`](.claude/skills/scenario-manager) | Builds a live Base/Bull/Bear switching layer with a selector and CHOOSE or INDEX assumption pulls — visible on the sheet, unlike Excel's hidden Scenario Manager. |
| [`sensitivity-tables`](.claude/skills/sensitivity-tables) | Builds one- and two-variable sensitivity grids on a key output, so you can show how NPV, IRR, EPS or margin moves as one or two drivers change. The grid values are computed in code; a native Excel Data Table is an optional step you add in Excel. |
| [`sensitivity-tornado`](.claude/skills/sensitivity-tornado) | Swings each input over a defined range, ranks the drivers by absolute impact and specifies the tornado chart that answers "which assumption actually matters?". |
| [`formula-audit-checker`](.claude/skills/formula-audit-checker) | Runs the pre-submission audit — trace precedents and dependents, hardcodes buried in formulas, circular references, error handling that masks problems — and returns a prioritised fix list. |
| [`output-summary-tab`](.claude/skills/output-summary-tab) | Designs the one-page executive summary tab: KPI tiles, the so-what narrative, a compact sensitivity table and the key assumptions, sized for one landscape page. |

## Install

Each skill is a folder containing a `SKILL.md` with YAML frontmatter, following the [Agent Skills](https://code.claude.com/docs/en/skills) format.

**Claude Code** — clone into your skills directory:

```bash
git clone https://github.com/plwanga-coder/andreworia.git
cp -r andreworia/.claude/skills/* ~/.claude/skills/
```

**Claude app (web, desktop, mobile)** — go to Settings → Capabilities → Skills and upload a skill folder, or zip one and upload it.

**A single skill** — copy just the folder you want. Each one is self-contained.

**This repo** — skills in `.claude/skills/` load automatically in any Claude Code session opened on the repo.

**Dependencies for the `.xlsx`-building skills** (`revenue-build`, `unit-economics`, `scenario-manager`, `sensitivity-tables`):

```bash
pip install openpyxl numpy
apt-get install libreoffice-calc   # libreoffice-core alone cannot open spreadsheets
```

## Usage

Skills load themselves when the conversation matches their description. You can also name one directly:

```
Use the model-architecture-template skill. I'm starting a three-year
operating model for a B2B SaaS business.
```

```
Use data-cleaning-for-excel on this pasted CRM export, then pivot it by region and month.
```

A natural chain for a full model:

```
data-cleaning-for-excel  →  model-architecture-template  →  inputs-calcs-outputs-design
                         →  revenue-build  →  unit-economics
                         →  scenario-manager  →  sensitivity-tables  →  sensitivity-tornado
                         →  formula-audit-checker  →  assumption-registry-builder
                         →  output-summary-tab
```

Clean the data first, agree the structure, build the drivers, then stress the model and audit it before anyone else sees it. The register and the summary tab are what you actually hand over.

## Contributing

Issues and pull requests are welcome. If a skill is wrong, or you have a sharper version, open a PR.

## Credits

The skills were originally published by Oria at [andreworia/claude-excel-skills](https://github.com/andreworia/claude-excel-skills) under the MIT licence. This repo carries audited and patched versions of them.

## License

[MIT](LICENSE) — use them commercially, modify them, redistribute them.
