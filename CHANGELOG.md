# xlsx-mini

## 0.3.0

### Minor Changes

- 8811f00: `streamXlsx` writes a workbook as a web `ReadableStream` of the file's bytes, produced as the consumer reads them: a `Response` body, a file, or `for await` over the chunks. Rows may be an async iterable, so a database cursor can feed a sheet page by page, and peak memory stays flat however many rows go out. Text is written inline by default; `strings: "shared"` keeps the shared-strings table instead. The streamed package uses data descriptors and, where the runtime has no raw deflate, stored deflate blocks, so sequential zip readers can open it too. `writeXlsx` is unchanged, byte for byte, and now reports `ASYNC_ROWS` instead of a `TypeError` when given an async iterable.

## 0.2.0

### Minor Changes

- 9d792d9: Styles grow beyond bold and number formats: `italic`, `fontSize`, `fontColor`, a solid `fill`, a `border` for all sides or per side, `align`, `verticalAlign`, `wrap` and `indent`, with colours as `#rrggbb`. A sheet takes `merges` as A1 ranges and `rowHeights` in points by row number. A cell object with a style and a `null` value now writes an empty styled cell, so a border or fill can run across a blank. Three new error codes: `INVALID_STYLE`, `INVALID_MERGE` and `INVALID_ROW_HEIGHT`. A style error on a cell now reports the property in its path, such as `sheets[0].rows[1][0].style.numberFormat`. Workbooks that use none of this render byte for byte as before.

## 0.1.0

### Minor Changes

- bb97199: First documented release. `writeXlsx` turns sheets of rows into `.xlsx` bytes with no runtime dependencies, on Node, Bun, Deno, Cloudflare Workers, Vercel's edge runtime and browsers: strings, numbers, booleans, `Date` and Temporal values, bold and number formats, column widths, frozen panes and autofilters, several sheets, `sheetFromObjects` for arrays of objects, a `coerce` hook and a type augmentation point for custom values, and `downloadXlsx` from `xlsx-mini/browser`. Every error is an `XlsxMiniError` with a code and a path. The golden fixtures open clean in Excel, Numbers, Google Sheets and LibreOffice.

## 0.0.1

### Patch Changes

- b40624d: Placeholder release that proves the release pipeline end to end. The first documented release is 0.1.0.
