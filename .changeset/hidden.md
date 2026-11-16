---
"xlsx-mini": minor
---

A column, a row and a sheet can be hidden: `hidden: true` on a column, on a row object or on a sheet, and on a column of `sheetFromObjects`. What is hidden stays in the file, so formulas and dropdowns read it, and the reader can show it again. A workbook opens on its first sheet that is not hidden; one whose sheets are all hidden fails with `NO_VISIBLE_SHEET`. A picture in a hidden row or column folds away with it.
