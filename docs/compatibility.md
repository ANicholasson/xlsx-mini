# Compatibility log

Which applications opened each golden fixture without a repair prompt, and how that was checked.
A row is only "clean" when the application opened the file without any dialog and the cell values
read back as written. Add a row per application and fixture whenever a fixture changes.

## minimal.xlsx (one sheet, strings and numbers)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                         | Date       |
| ---------------- | -------- | -------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; all 15 cells read back through AppleScript (`value of range "A1:C5"`), including escaped text, emoji, padded whitespace and `1e+21` | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; imported table read back through AppleScript cell by cell                                                                           | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; all 15 cells displayed as written, `1E+21` shown for the large number                                                        | 2026-11-04 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | `soffice --headless --convert-to csv`; the CSV matches every cell (empty cells empty, whitespace kept, `1E+021` for the large number)                      | 2026-11-04 |

Independent zip checks on the same fixture: `unzip -t` (Info-ZIP 6.00) and Python's `zipfile` both report no errors.

Note for the headless conversion: with the sheet-selection token set to `-1` in the CSV filter options, LibreOffice writes one file per sheet named `<fixture>-<sheet>.csv`; without filter options it writes the first sheet to `<fixture>.csv`.

## cell-types.xlsx (one row per kind of value: text, number, booleans, Date, the four Temporal types, a wrapped value, empty, stripped control characters, a kept newline, a coerced bigint)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                         | Date       |
| ---------------- | -------- | -------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; all 30 cells read back through AppleScript; date cells carry the formats `m/d/yy h:mm` and `m/d/yy`                                 | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; imported table read back cell by cell, dates shown as dates, booleans as TRUE/FALSE, newline kept                                   | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; no repair notice, every value displayed as written: dates as dates, booleans as TRUE/FALSE, the newline kept inside its cell | 2026-11-06 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless CSV conversion matches every cell, dates rendered from their formats                                                                              | 2026-11-04 |

## minimal-deflate.xlsx (the minimal workbook with DEFLATE entries; compared part by part, not byte by byte, because compressor output differs by platform)

| Application      | Version  | Platform | Result | How it was checked                                                                     | Date       |
| ---------------- | -------- | -------- | ------ | -------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; all 15 cells read back through AppleScript                      | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; imported table read back cell by cell                           | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; no repair notice, the same table as the minimal workbook | 2026-11-06 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless CSV conversion matches every cell                                             | 2026-11-04 |

Independent inflate checks on the same fixture: `unzip -t` and Python's `zipfile` both report no errors.

## multi-sheet.xlsx (three sheets, names with `&` and `<>`, text shared across sheets, one empty sheet)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                           | Date       |
| ---------------- | -------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; tab names and order read back through AppleScript, cells of every sheet read back                                                                                                     | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; sheet names and order read back, sample cells read back                                                                                                                               | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; three tabs in order with `&` and `<>` intact in their names, the empty sheet present, the cells of each sheet displayed as written on a visual check, no cell-by-cell readback | 2026-11-06 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless CSV conversion with the all-sheets token wrote one file per sheet, named after the sheet, each matching                                                                                             | 2026-11-04 |

## styles.xlsx (bold header and total row, a kroner currency column, a percent column, an ISO-date column, a default-formatted date column)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                      | Date       |
| ---------------- | -------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; every cell's value, number format and bold flag read back through AppleScript and matched the input                                                                              | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; formatted values read back: `1 234,50 kr`, `10,0 %`, `2026-11-01`, `11/3/26 14:30`                                                                                               | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; bold header and total row, the kroner, percent and ISO-date columns displayed with their formats, values displayed as written on a visual check, no cell-by-cell readback | 2026-11-06 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless HTML conversion shows `1,234.50 kr`, `10.0%`, ISO dates and bold header and total rows; its CSV exporter prints the currency column raw, a quirk of that filter, so format checks use HTML     | 2026-11-04 |

## layout.xlsx (column widths 12, 32 and 14, a frozen header row and first column, an autofilter on the first sheet only, an unfiltered second sheet)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                                                                                                 | Date       |
| ---------------- | -------- | -------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; AppleScript reports autofilter on sheet 1 and off on sheet 2, freeze panes with split row 1 and split column 1, column widths 11.17, 31.17 and 13.17 (Excel reports a width minus its 0.83 padding, as for its own files), default 10 for the fourth column | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; both sheets present, widths imported in proportion (72, 192 and 84 points), values and currency format intact. Numbers does not turn frozen panes into header rows on import; that is its behaviour for any `.xlsx`                                         | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; column widths, the frozen header row and first column, filter controls on the first sheet only, both sheets' cells displayed as written on a visual check, no cell-by-cell readback                                                                  | 2026-11-06 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless CSV per sheet matches; a headless round trip to `.xlsx` keeps the frozen pane, the three column widths and the autofilter range, so all three were parsed                                                                                                                 | 2026-11-04 |

## orders.xlsx (built with `sheetFromObjects`: bold header, widths, a kroner column, a date-time column, frozen header row, autofilter)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                              | Date       |
| ---------------- | -------- | -------- | ------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean  | Opened from the shell; AppleScript reports the bold header, the `#,##0.00 "kr"` and `yyyy-mm-dd hh:mm` formats, freeze at row 1, autofilter on, and the four widths                                             | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; formatted values read back: `1 234,50 kr`, `2026-11-01 09:30`                                                                                                                            | 2026-11-04 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; bold header, widths, the kroner and date-time columns formatted, frozen header row and filter controls, the rows displayed as written on a visual check, no cell-by-cell readback | 2026-11-06 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless HTML conversion shows the bold header, `1,234.50 kr` and `2026-11-01 09:30`                                                                                                                            | 2026-11-04 |

## appearance.xlsx (a title merged across four columns with a fill, white bold 16-point text and centred alignment; a filled and thin-bordered header row; a wrapped top-aligned cell; an indented cell; a total row ruled medium above and thick below, including a styled empty cell; row heights of 28 and 48 points; an italic grey note)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                                                                                                                            | Date       |
| ---------------- | -------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.4 | macOS    | clean  | Opened from the shell; read back through AppleScript: merge area `$A$1:$D$1`, bold 16-point white font on fill 31,78,121, centred; header fill 217,225,242 with a continuous top border; wrap and top alignment; indent level 1; thick bottom rule on the total row, also on the empty cell; italic grey note | 2026-11-06 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; read back through AppleScript: title background and white text at 16 points centred, header background, wrap and top alignment, grey note, `9 998,00 kr` formatted, row heights 28 and 48 as written                                                                                   | 2026-11-06 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; the title merged across four columns on its fill, the header fill and borders, the wrapped description, the ruled total row, the italic grey note and the kroner amounts all displayed as written on a visual check, no cell-by-cell readback; dated 2026-11-06                 |            |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless CSV conversion matches every cell; the HTML conversion shows the title spanning four columns, both fill colours, the grey font colour, eight bold cells, one italic cell and the kroner amounts                                                                                                      | 2026-11-06 |

## Runtimes

Every host writes the same workbook (`test/runtimes/scenario.mjs`) through the built package and
reports back. The STORE output must be byte-identical on every host and to the source run in the
test's own Node; the default output must decompress part by part to the same XML. The matrix runs
inside `pnpm test`; this table records what each host did with compression and how it was loaded.

| Host                | Version                                                                    | Loaded through                                                        | Default output                                                                                   | Date       |
| ------------------- | -------------------------------------------------------------------------- | --------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ | ---------- |
| Node, ESM           | 24.22.1                                                                    | `import` by package name through the exports map                      | DEFLATE                                                                                          | 2026-11-05 |
| Node, CommonJS      | 24.22.1                                                                    | `require` of the ESM-only package and of the ESM scenario             | DEFLATE                                                                                          | 2026-11-05 |
| Bun                 | 1.4.4                                                                      | `import` by package name through the exports map                      | DEFLATE                                                                                          | 2026-11-05 |
| Deno                | 2.9.6                                                                      | `import` by package name through the exports map                      | DEFLATE                                                                                          | 2026-11-05 |
| Cloudflare workerd  | 1.20260730.1 through Miniflare 4.20260730.0, compatibility date 2026-07-01 | rolldown bundle of `dist/` as a module worker, called through `fetch` | DEFLATE: raw deflate confirmed in workerd                                                        | 2026-11-05 |
| Vercel edge runtime | `@edge-runtime/vm` 5.0.0                                                   | rolldown bundle of `dist/` as a classic script, evaluated in the VM   | STORE: the VM has no `CompressionStream`, so the writer fell back; the parts are still identical | 2026-11-05 |

The edge VM package was last published in December 2024 and may lag the hosted runtime; its row
proves the fallback path, not deflate on Vercel. Both entries load on every host, and
`downloadXlsx` reports `NO_DOM` on each of them.

## streamed.xlsx (the appearance workbook written by `streamXlsx` with stored deflate blocks and inline strings: data descriptors on every entry, the sheet before the workbook, styles and shared-strings parts, no dimension element; also the same workbook streamed with raw deflate, which is what a server sends by default)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                                                        | Date       |
| ---------------- | -------- | -------- | ------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.4 | macOS    | clean  | Both files opened from the shell without a prompt; read back through AppleScript: merge area `$A$1:$D$1`, bold header, row heights identical to those read from appearance.xlsx, the indented description and the empty styled total cell | 2026-11-06 |
| Apple Numbers    | 14.4     | macOS    | clean  | Both files opened from the shell; title, description and the empty total cell read back through AppleScript as written                                                                                                                    | 2026-11-06 |
| Google Sheets    | web      | browser  | clean  | Both files uploaded to Drive and opened; the workbook displayed as written on a visual check, no cell-by-cell readback                                                                                                                    | 2026-11-09 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | Headless CSV conversion of both files is identical to the CSV recorded for appearance.xlsx; `unzip -t` accepts both                                                                                                                       | 2026-11-06 |

## Browser download (`downloadXlsx` from `xlsx-mini/browser`)

| Browser       | Version       | Platform | Result | How it was checked                                                                                                                                                                                                               | Date       |
| ------------- | ------------- | -------- | ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Google Chrome | 155.0.8059.86 | macOS    | clean  | A page bundling `dist/` wrote a four-row workbook and called `downloadXlsx` on load in a fresh profile; the file landed in Downloads within three seconds, `unzip -t` and Python's `zipfile` accept it, and Excel opens it clean | 2026-11-05 |

## formulas.xlsx (per-row products and a sum without cached values; a sum, a count, a comparison and a text formula with cached values; text that starts with `=`)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                         | Date       |
| ---------------- | -------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ---------- |
| Microsoft Excel  | 16.114.4 | macOS    | clean  | Opened from the shell; read back through AppleScript: D2 holds a formula and shows 9998, D5 computes 15545, B8 `TRUE`, B9 `Items: 3`; B10 shows `=SUM(B2:B4)` as text and holds no formula | 2026-11-09 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell; read back through AppleScript: D5 15545, B8 `true`, B9 `Items: 3`, B10 the text `=SUM(B2:B4)`; D2 reports the formula `B2×C2`                                       | 2026-11-09 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; no dialog, every sheet shown as written                                                                                                                      | 2026-11-09 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | `soffice --headless --convert-to csv`; the CSV carries the computed results (9998, 5198, 349, 15545, 7, 3, `TRUE`, `Items: 3`) and the text `=SUM(B2:B4)`                                  | 2026-11-09 |

## hyperlinks.xlsx (two external links with a query string and a `mailto:`, a jump to another sheet, a jump within the sheet, a jump back to a range; underlined blue text)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                                                                       | Date       |
| ---------------- | -------- | -------- | ------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.4 | macOS    | clean  | Opened from the shell; read back through AppleScript: four hyperlinks on the first sheet, addresses `https://example.com/docs?a=1&b=2` and `mailto:hello@example.com`, sub-addresses `'Summary'!B2` and `'Links'!A1`, single underline on the link cells | 2026-11-09 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell without a dialog; both sheets show the link text as blue underlined links                                                                                                                                                          | 2026-11-09 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; no dialog, every sheet shown as written                                                                                                                                                                                    | 2026-11-09 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | `soffice --headless --convert-to csv`; the CSV of both sheets matches the link text cell for cell                                                                                                                                                        | 2026-11-09 |

## validation.xlsx (a dropdown from a literal list with an input message, a dropdown from a range on another sheet, whole-number bounds with an error message, date bounds with a warning and no blanks, a whole-column decimal bound with an information message)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                                                                                     | Date       |
| ---------------- | -------- | -------- | ------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.4 | macOS    | clean  | Opened from the shell; read back through AppleScript: list source `Open,Closed,On hold` with its prompt, range source `=Lists!$A$2:$A$4`, bounds 1 and 1000 with their error text, warning style and no blanks on the date rule, the whole-column rule present at F500 | 2026-11-09 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell without a dialog; both sheets show every value as written                                                                                                                                                                                        | 2026-11-09 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; no dialog, every sheet shown as written                                                                                                                                                                                                  | 2026-11-09 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | `soffice --headless --convert-to csv`; the CSV of both sheets matches cell for cell                                                                                                                                                                                    | 2026-11-09 |

## print.xlsx (landscape A4 fitted to one page wide, half-inch side margins, a header and a footer with page codes, the first row repeated on every page)

| Application      | Version  | Platform | Result | How it was checked                                                                                                                                                                                                                                      | Date       |
| ---------------- | -------- | -------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.4 | macOS    | clean  | Opened from the shell; page setup read back through AppleScript: landscape, fit to 1 page wide and as many tall, left margin 36 points and top 54 points, header `Quarterly report` and `&D`, footer `&A` and `Page &P of &N`, print title rows `$1:$1` | 2026-11-09 |
| Apple Numbers    | 14.4     | macOS    | clean  | Opened from the shell without a dialog; all 31 rows shown with the currency format                                                                                                                                                                      | 2026-11-09 |
| Google Sheets    | web      | browser  | clean  | Uploaded to Drive and opened; no dialog, every sheet shown as written                                                                                                                                                                                   | 2026-11-09 |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean  | `soffice --headless --convert-to csv`; the CSV matches cell for cell                                                                                                                                                                                    | 2026-11-09 |
