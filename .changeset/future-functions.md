---
"xlsx-mini": patch
---

Formulas with functions newer than Excel 2007 work as typed. `XLOOKUP`, `IFS`, `TEXTJOIN`, `CONCAT`, `MAXIFS`, `STDEV.S` and the rest of Excel's later functions are now stored under the prefix the file format requires; before, such a formula showed `#NAME?` in Excel and LibreOffice until its cell was edited. A formula that calls `LET` or `LAMBDA` is written exactly as given, and a function that returns an array shows its first value rather than spilling. Google Sheets does not know the name Excel stores `FILTER` and `SORT` under, so those two show `#NAME?` there.
