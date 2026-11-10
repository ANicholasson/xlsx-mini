---
"xlsx-mini": minor
---

Auto column width. `width: "auto"` on a column fits it to its widest cell, and `autoWidth: true` on a sheet, or in `sheetFromObjects`, does that for every column without a numeric width. The width is estimated from what each cell shows, text character by character and numbers and dates by their format, bold a little wider, from 3 to 60 characters, with no font measured. On `writeXlsx` every row counts; on `streamXlsx` the first 1 000 rows of a sheet set the widths. A workbook that uses neither renders byte for byte as before.
