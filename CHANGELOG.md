# xlsx-mini

## 0.6.0

### Minor Changes

- cfa1ec8: Images. A cell `{ image: bytes }` holds a PNG, anchored to the cell at its pixel size or at a given `width` or `height`, moving with the cell without stretching. The row grows to the picture, and so does a column without a numeric width; identical images are written once. A `sheetFromObjects` value function may now return any cell, an image included. New error code `INVALID_IMAGE`. A streamed package now writes its content types last, since only the sheets know whether a drawing part exists; a workbook without images writes the same bytes as before on `writeXlsx`.

## 0.5.0

### Minor Changes

- a4437a3: Auto column width. `width: "auto"` on a column fits it to its widest cell, and `autoWidth: true` on a sheet, or in `sheetFromObjects`, does that for every column without a numeric width. The width is estimated from what each cell shows, text character by character and numbers and dates by their format, bold a little wider, from 3 to 60 characters, with no font measured. On `writeXlsx` every row counts; on `streamXlsx` the first 1 000 rows of a sheet set the widths. A workbook that uses neither renders byte for byte as before.

## 0.4.0

### Minor Changes

- b075101: Formulas, hyperlinks, data validation and print setup, on both `writeXlsx` and `streamXlsx`.

  - A cell `{ formula: "SUM(B2:B4)" }`, with an optional cached `value`; formulas are escaped, never evaluated, and text starting with `=` stays text. A workbook with formulas asks Excel to calculate on open.
  - A cell `{ value, href }` links to an absolute URL; `{ value, cell }` and `{ value, sheet, cell }` jump within the workbook. `underline` joins `CellStyle`.
  - A sheet's `validations`: dropdowns from a list or a range, number and date bounds, input and error messages, on a cell, a range or whole columns.
  - A sheet's `print`: orientation, paper size, fit to pages, margins, header and footer with Excel's codes, repeated header rows.
  - New error codes `INVALID_FORMULA`, `INVALID_HYPERLINK`, `INVALID_VALIDATION` and `INVALID_PRINT_SETUP`.

  A workbook that uses none of these writes the same bytes as before.

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
