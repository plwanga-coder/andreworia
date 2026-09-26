---
name: data-cleaning-for-excel
description: Turns pasted or exported data with mixed formats, blank rows, or duplicates into a clean, consistently formatted range ready for analysis. Use it right after pasting raw data into a worksheet.
---

# Data Cleaning for Excel

## When to use
Use this immediately after pasting or importing raw data, a CSV export, a copy from another
system, or a table copied out of a document, before you build a pivot table or a formula on
top of it.

## Instructions
1. **Never clean in place.** Keep the original on a `Raw` tab, untouched. Write the cleaned range to a `Clean` tab and the change log to a `Change Log` tab. The raw data is the audit trail.
2. **Profile each column first.** For every column, report the inferred type, count of blanks, count of distinct values and a few example values. Flag columns where types are mixed (text numbers next to real numbers, dates next to text).
3. **Fix the structure.** Keep exactly one header row. Remove fully blank rows, repeated header rows (common in multi-page exports) and subtotal or total rows mixed into the data. Unmerge merged cells and fill the value down where a merge meant "same as above".
4. **Numbers.** Convert text numbers to real numbers, handling: thousands separators, currency symbols, `%` signs (divide by 100), parentheses negatives `(1,234)`, trailing minus `1234-`, and European decimal commas when the whole column uses them. Do NOT convert identifier-like columns (account numbers, ZIP/postcodes, phone numbers, IDs with leading zeros); keep those as text.
5. **Dates.** Convert text dates to real Excel dates in a single format. Detect ambiguous day/month order: if any value in the column has a first part greater than 12, the column is day-first; if none does, the order is ambiguous, so ask rather than guess. Convert Excel serial numbers and ISO strings too. Flag impossible dates (31-Feb) and far-out-of-range years.
6. **Text.** Trim leading, trailing and repeated spaces, including non-breaking spaces (CHAR(160)) and line breaks. Normalize casing for category fields only. Map label variants that clearly mean the same thing ("N. America", "North America ", "NA") to one value, and list every mapping in the change log. Ask when a mapping is not obvious.
7. **Blanks vs zero.** Distinguish a true zero from a missing value. Leave missing values blank (not 0) unless the user confirms that blank means zero. Report the count per column.
8. **Duplicates.** State the rule that defines a duplicate: an exact full-row match, or a match on a business key (e.g. Invoice ID). Remove exact duplicates. For key duplicates whose other fields differ, flag them rather than delete them.
9. **Report what changed**, row by row where practical: row reference, column, old value, new value, rule applied. Finish with a summary: rows in, rows out, and counts per rule.

## Implementation notes
- When generating a .xlsx with code, use pandas for the cleaning and openpyxl to write the three tabs. Set explicit number formats on converted columns (dates `yyyy-mm-dd` or the user's locale, numbers `#,##0.00`) so Excel does not re-interpret them. Write ID columns with the text format `@`.
- When working directly in Excel, the formula equivalents are `TRIM(SUBSTITUTE(A2,CHAR(160)," "))`, `VALUE`/`NUMBERVALUE(A2,",",".")` and `DATEVALUE`, or Power Query for repeatable cleans.

## Example prompts
- "Use the data-cleaning-for-excel skill on this pasted CSV before I pivot it."
- "Clean up this exported customer list, the dates are in three different formats."

## Example
A 1,204-row CRM export has a repeated header on every 50th row, dates as "03/04/2024", "2024-04-05" and "5 Apr 2024", amounts as "$1,250.00" and "(300)", and region values "EMEA", "Emea " and "Europe/ME". Result: 24 repeated headers removed; the dates are day-first (values like "25/03/2024" exist), so all are converted to real dates; amounts are converted to numbers with "(300)" becoming -300; the three region labels are mapped to "EMEA"; 6 exact duplicates are removed and 2 Invoice-ID conflicts are flagged. The change log lists every edit: 1,204 rows in, 1,174 out.

## Output
A cleaned, consistently formatted range on its own tab, the untouched raw data, and a change log describing what was standardized, removed, converted or flagged, with a summary of counts.
