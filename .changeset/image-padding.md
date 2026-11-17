---
"xlsx-mini": minor
---

An image cell takes `padding`: the space between the picture and each edge of its cell, in pixels. `{ image, height: 40, padding: 10 }` gives a row of 60 pixels with the picture in its middle, so one picture is 20 pixels from the next down a column; a column without a numeric width grows the same way. Without `padding` a picture is placed as before, two pixels inside the top-left corner of its cell.
