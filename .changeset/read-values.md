---
"xlsx-mini": minor
---

A new entry, `xlsx-mini/read`, reads the values of an `.xlsx` file. `openXlsx` lists the sheets and reads each sheet's rows, `objectsFromRows` reads rows as objects by the column list `sheetFromObjects` takes, and where a part it reads is malformed, the file is refused with an `XlsxReadError` and a code. The entry is experimental, and any 0.x minor may change it. The README section "Reading (experimental)" says what it reads, where it runs and what it costs.
