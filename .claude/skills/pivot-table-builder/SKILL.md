---
name: pivot-table-builder
description: Specifies and builds a pivot table from a flat data range, choosing rows, columns, values, and filters. Use it when you have raw tabular data and need a summary view without building it cell by cell.
---

# Pivot Table Builder

## When to use
Use this when you have a flat data range, one row per record, and need a summarized view by
category, time period, or segment, and would rather describe the summary you want than
build the pivot table manually.

## Instructions
1. Confirm the data range is genuinely flat: one header row, no merged cells, no subtotal rows mixed into the data. If not, run data-cleaning-for-excel first.
2. Ask what question the summary needs to answer if it is not stated, since that decides which fields go where.
3. Propose the pivot table layout: which field goes to Rows, which to Columns, which to Values (and the aggregation, sum, count, or average), and any Filters. Put time periods on Columns and categories on Rows unless there are more periods than categories.
4. Build it with the method that fits where you are working (see "How to build").
5. Flag any field that will not aggregate cleanly, for example a text column dropped into Values by mistake, numbers stored as text (they count but do not sum), blanks in a category field, or daily dates that should be grouped into months.
6. Check the pivot: the grand total equals the total of the source column (for example `SUM` of the Amount column), so no rows were dropped by a filter or by bad types.

## How to build
- **Working directly in Excel (the user or an add-in operating Excel):** first convert the range to a Table (Ctrl+T) so the pivot expands with new rows. Then Insert > PivotTable, place it on a new sheet, and drag the fields as specified. Group dates by Months/Years (right-click > Group).
- **Generating a .xlsx with code (openpyxl):** openpyxl cannot create a native PivotTable. Pick one of these and tell the user which you used:
  - **Live formula grid (preferred):** write the unique row and column labels, then fill the body with `SUMIFS`/`COUNTIFS`/`AVERAGEIFS` pointing at the data tab. For example `=SUMIFS(Data!$D:$D,Data!$B:$B,$A5,Data!$C:$C,B$4)`, plus row and column totals with `SUM`. It recalculates when the data changes, though new category labels still need adding.
  - **Static summary:** compute with pandas `pivot_table(index=..., columns=..., values=..., aggfunc=..., margins=True)` and write it as values labelled "static snapshot".
  - Either way, include the steps for the user to insert a native PivotTable in Excel if they want slicers or drill-down.

## Example prompts
- "Use the pivot-table-builder skill to summarize this sales data by region and month."
- "Build a pivot showing headcount by department and level from this raw HR export."

## Example
Sales data with columns Date, Region, Product, Amount (2,400 rows). Question: "Which region grew fastest this year?" Layout: Rows = Region, Columns = Date grouped by month, Values = Sum of Amount, Filter = Year 2025. Built as a SUMIFS grid with the month start and end dates in the column headers: `=SUMIFS(Data!$D:$D,Data!$B:$B,$A5,Data!$A:$A,">="&B$3,Data!$A:$A,"<"&EDATE(B$3,1))`. Check: the grand total equals `=SUMIFS(Data!D:D,Data!A:A,">="&DATE(2025,1,1),Data!A:A,"<"&DATE(2026,1,1))`. Flag: 14 Amount values were text and were converted before summing.

## Output
A pivot table specification (rows, columns, values, filters), the built pivot or equivalent summary on the worksheet (stating whether it is native, formula-driven or static), the grand-total check, and a one-line note on what it answers.
