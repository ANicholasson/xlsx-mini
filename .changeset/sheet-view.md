---
"xlsx-mini": minor
---

A sheet can say how it is shown when it opens: `view: { gridlines: false, zoom: 150, rightToLeft: true }` hides the lines between cells, sets the zoom in percent and lays the sheet out from the right. It combines with `freeze`. A zoom outside 10 to 400 or a setting of the wrong type fails with `INVALID_VIEW`.
