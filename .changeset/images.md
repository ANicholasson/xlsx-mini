---
"xlsx-mini": minor
---

Images. A cell `{ image: bytes }` holds a PNG, anchored to the cell at its pixel size or at a given `width` or `height`, moving with the cell without stretching. The row grows to the picture, and so does a column without a numeric width; identical images are written once. A `sheetFromObjects` value function may now return any cell, an image included. New error code `INVALID_IMAGE`. A streamed package now writes its content types last, since only the sheets know whether a drawing part exists; a workbook without images writes the same bytes as before on `writeXlsx`.
