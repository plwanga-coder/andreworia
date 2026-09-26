# Skills Audit — claude-excel-skills

Audit date: 2026-09-26. Scope: all 12 `skills/*/SKILL.md` files (read in full), repo history (3 commits, single author), and assets.

## Verdict

**Safe to install.** The findings below are about technical accuracy, not security. Most skills are solid method guides. Five have errors that would give wrong results or break a build if followed exactly.

## Security review

| Check | Result |
|:--|:--|
| Executable content (scripts, binaries) in skill folders | None. Each folder holds only `SKILL.md` |
| Network access, URLs, curl/wget, external endpoints | None in any skill |
| Credential or secret handling, exfiltration patterns | None |
| Prompt injection ("ignore previous", system-prompt overrides, hidden HTML comments) | None |
| Hidden or invisible Unicode (zero-width, bidi, tag chars) | None. All skill files are pure ASCII |
| Shell commands the skills tell Claude to run | Only `libreoffice --headless --convert-to xlsx`, to recalculate locally. Low risk |
| Frontmatter validity (`name` matches folder, `description` present) | 12/12 pass |
| Assets (`.github/assets`) | PNG/GIF images used only by the README |
| Licence | MIT |

Note: the README is mostly marketing for Oria. The skills themselves contain no promotional or tracking content.

## Inventory

| Skill | Type | Size | Quality |
|:--|:--|--:|:--|
| model-architecture-template | Method | 11.4 KB | Good. 1 formula bug |
| inputs-calcs-outputs-design | Method | 9.8 KB | Good |
| assumption-registry-builder | Method | 11.2 KB | Good. 2 inconsistencies |
| data-cleaning-for-excel | Method | 1.4 KB | Thin but correct |
| pivot-table-builder | Method | 1.4 KB | Thin. Build step not specified |
| revenue-build | Builds .xlsx | 7.0 KB | Good. Inherits Data Table issue |
| unit-economics | Builds .xlsx | 7.3 KB | Good. 1 check too strict |
| scenario-manager | Builds .xlsx | 7.7 KB | Good |
| sensitivity-tables | Builds .xlsx | 7.9 KB | **Data Table approach will not work as written** |
| sensitivity-tornado | Method | 9.3 KB | Good. Example table mis-sorted |
| formula-audit-checker | Method | 10.4 KB | Good. 2 technical inaccuracies |
| output-summary-tab | Method | 10.2 KB | Good |

## Technical findings

### High

1. **sensitivity-tables / revenue-build: native Data Tables cannot be built the way these skills describe.**
   - Excel requires a Data Table's row and column input cells to be on the **same sheet** as the table. Both skills put the table on a `Sensitivity` tab and point the input cells at `Model!B6` / `Drivers!...`. Excel rejects that with "Input cell reference is not valid".
   - openpyxl cannot write a real `{=TABLE()}` array. Written as a normal formula, `=TABLE(...)` shows as `#NAME?` in both Excel and LibreOffice (LibreOffice uses `MULTIPLE.OPERATIONS`). That conflicts with the skills' own rule to "verify zero #NAME? errors".
   - **Workaround:** use the code-computed grid that sensitivity-tables already describes as its fallback. Or put the table on the same sheet as the driver cells, or use `MULTIPLE.OPERATIONS` when targeting LibreOffice.

### Medium

2. **model-architecture-template: the Checks-tab traffic light misses some failures.** `COUNTIF(checks_range,"FAIL")` only matches the exact text `FAIL`. The skill's own checks return strings like `"FAIL -- WACC must exceed terminal growth"`, which it won't count. **Fix:** `COUNTIF(checks_range,"FAIL*")`.
3. **formula-audit-checker: Go To Special > Constants does not find hardcodes inside formulas.** It only selects cells that hold nothing but a constant. A cell like `=A1*0.25` isn't picked up. Use Find in Formulas, `FORMULATEXT`, or Inquire / a script to find embedded constants.
4. **formula-audit-checker: wrapping every lookup in IFERROR is risky.** `IFERROR(...,0)` hides real breakages, which works against what an audit is for. Prefer `IFNA`, or return an explicit flag string that the Checks tab counts.
5. **assumption-registry-builder: RANK runs on signed impact.** Negative impacts rank as "least impactful". Rank on `ABS()` of the swing, or on the absolute range column. The example also puts text (`+$18M / -$14M`) in the column that RANK reads.

### Low

6. **assumption-registry-builder:** the High/Medium/Low thresholds are `<=3 / <=7` in the method but `<=2 / <=5` in the example.
7. **sensitivity-tornado:** the example table says it is "sorted by absolute impact" but lists WACC ($68M) above Revenue CAGR ($77M).
8. **unit-economics:** the Checks tab requires closed-form LTV and discounted finite-horizon cohort LTV to agree within a tolerance. They differ by design (discounting and truncation), so the tolerance needs to be generous or the check always fails. There is also a confusing inline note, "`'Unit Economics'!B6` lifetime churn"; B6 is retention. The final formula is correct.
9. **formula-audit-checker:** `Ctrl+[` selects precedents; it doesn't draw Trace Precedents arrows.
10. **pivot-table-builder:** says "build the pivot table" but gives no method. openpyxl cannot create pivot tables, so in the Claude app this ends up as a pandas `pivot_table` summary written as a static range.
11. **data-cleaning-for-excel / pivot-table-builder:** very short compared with the rest (no method detail or worked example).

Worked examples checked and arithmetically correct: revenue-build (1,000 + 200 − 30 = 1,170), unit-economics (CAC 500, payback 12.5 mo, LTV 1,000, LTV/CAC 2.0x, retention 100% / 96% / 92.2%), scenario-manager (Bull delta 380).

## Runtime dependencies (for the four .xlsx-building skills)

- `openpyxl` (and `numpy` for sensitivity-tables): **not installed in this container**. Install with `pip install openpyxl numpy` before use.
- `libreoffice` (headless recalc): present.

## Installation performed

- **Project level:** copied to `.claude/skills/<name>/SKILL.md` in this repo, unmodified. They load automatically in any Claude Code session opened on this repo.
- **User level (this session):** copied to `~/.claude/skills/`. This cloud container is temporary, so copy them to `~/.claude/skills/` on your own machine too, or upload each folder in the Claude app (Settings → Capabilities → Skills).
- The installed copies match `skills/` byte for byte (`diff -r` clean). The findings above were **not** patched in, so the skills stay identical to upstream.
