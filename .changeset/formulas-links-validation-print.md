---
"xlsx-mini": minor
---

Formulas, hyperlinks, data validation and print setup, on both `writeXlsx` and `streamXlsx`.

- A cell `{ formula: "SUM(B2:B4)" }`, with an optional cached `value`; formulas are escaped, never evaluated, and text starting with `=` stays text. A workbook with formulas asks Excel to calculate on open.
- A cell `{ value, href }` links to an absolute URL; `{ value, cell }` and `{ value, sheet, cell }` jump within the workbook. `underline` joins `CellStyle`.
- A sheet's `validations`: dropdowns from a list or a range, number and date bounds, input and error messages, on a cell, a range or whole columns.
- A sheet's `print`: orientation, paper size, fit to pages, margins, header and footer with Excel's codes, repeated header rows.
- New error codes `INVALID_FORMULA`, `INVALID_HYPERLINK`, `INVALID_VALIDATION` and `INVALID_PRINT_SETUP`.

A workbook that uses none of these writes the same bytes as before.
