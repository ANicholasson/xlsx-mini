---
"xlsx-mini": minor
---

Styles grow beyond bold and number formats: `italic`, `fontSize`, `fontColor`, a solid `fill`, a `border` for all sides or per side, `align`, `verticalAlign`, `wrap` and `indent`, with colours as `#rrggbb`. A sheet takes `merges` as A1 ranges and `rowHeights` in points by row number. A cell object with a style and a `null` value now writes an empty styled cell, so a border or fill can run across a blank. Three new error codes: `INVALID_STYLE`, `INVALID_MERGE` and `INVALID_ROW_HEIGHT`. A style error on a cell now reports the property in its path, such as `sheets[0].rows[1][0].style.numberFormat`. Workbooks that use none of this render byte for byte as before.
