---
"xlsx-mini": patch
---

Row heights show as written in Excel for Mac. A sheet without a frozen pane or a `view` was written without a sheet view, and Excel for Mac then showed every row that has a height a third taller than asked: 40 points for 30. A picture kept its size, so it no longer filled its row. Every sheet now carries a sheet view. Numbers and LibreOffice showed the heights right before and still do.
