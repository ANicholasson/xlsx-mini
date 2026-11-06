# xlsx-mini

## 0.1.0

### Minor Changes

- bb97199: First documented release. `writeXlsx` turns sheets of rows into `.xlsx` bytes with no runtime dependencies, on Node, Bun, Deno, Cloudflare Workers, Vercel's edge runtime and browsers: strings, numbers, booleans, `Date` and Temporal values, bold and number formats, column widths, frozen panes and autofilters, several sheets, `sheetFromObjects` for arrays of objects, a `coerce` hook and a type augmentation point for custom values, and `downloadXlsx` from `xlsx-mini/browser`. Every error is an `XlsxMiniError` with a code and a path. The golden fixtures open clean in Excel, Numbers, Google Sheets and LibreOffice.

## 0.0.1

### Patch Changes

- b40624d: Placeholder release that proves the release pipeline end to end. The first documented release is 0.1.0.
