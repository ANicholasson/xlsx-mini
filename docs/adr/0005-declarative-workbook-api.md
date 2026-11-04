# A declarative, pure writeXlsx(workbook, options) instead of a builder

The caller passes one plain workbook object (sheets with rows of cells) and receives bytes. Rows are arrays of cells; a cell is a primitive or `{ value, style }`; style is a nested object so the cell type stays stable as borders and fills arrive; columns carry default styles and cell style wins; `rows` is any iterable so a cursor or generator can feed a sheet. The cell type is inferred from the JavaScript type, so there is no `type` field. Data (`Workbook`) and how to write it (`WriteOptions`: compression, coerce) are separate arguments so the data object stays serialisable. A convenience `sheetFromObjects` covers the array-of-objects case with opinionated defaults (bold, frozen header row, autofilter).

## Considered options

- Builder class in the ExcelJS style: familiar but stateful and a larger surface.
- Constructor that writes and downloads in one step: couples the writer to the browser.
