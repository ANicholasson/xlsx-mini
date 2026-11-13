---
"xlsx-mini": minor
---

A row can carry a style and a height. A row is an array of cells as before, or an object `{ cells, style, height }`: the style reaches every position in `cells`, the empty ones included, so a header, a band or a total row takes one style instead of one per cell, and a cell's own style still wins, property by property. `sheetFromObjects` takes `rowStyle` to style a row from its object. The type `Row` is now a union, so code that reads `row.length` from a `Row` narrows it first, and a row that is neither form fails with `INVALID_ROW`.
