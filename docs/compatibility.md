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

| Application      | Version  | Platform | Result  | How it was checked                                                                                                         | Date       |
| ---------------- | -------- | -------- | ------- | -------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean   | Opened from the shell; all 30 cells read back through AppleScript; date cells carry the formats `m/d/yy h:mm` and `m/d/yy` | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean   | Opened from the shell; imported table read back cell by cell, dates shown as dates, booleans as TRUE/FALSE, newline kept   | 2026-11-04 |
| Google Sheets    | web      | browser  | pending | Needs a manual upload and a visual check                                                                                   |            |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean   | Headless CSV conversion matches every cell, dates rendered from their formats                                              | 2026-11-04 |

## minimal-deflate.xlsx (the minimal workbook with DEFLATE entries; compared part by part, not byte by byte, because compressor output differs by platform)

| Application      | Version  | Platform | Result  | How it was checked                                                | Date       |
| ---------------- | -------- | -------- | ------- | ----------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean   | Opened from the shell; all 15 cells read back through AppleScript | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean   | Opened from the shell; imported table read back cell by cell      | 2026-11-04 |
| Google Sheets    | web      | browser  | pending | Needs a manual upload and a visual check                          |            |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean   | Headless CSV conversion matches every cell                        | 2026-11-04 |

Independent inflate checks on the same fixture: `unzip -t` and Python's `zipfile` both report no errors.

## multi-sheet.xlsx (three sheets, names with `&` and `<>`, text shared across sheets, one empty sheet)

| Application      | Version  | Platform | Result  | How it was checked                                                                                               | Date       |
| ---------------- | -------- | -------- | ------- | ---------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean   | Opened from the shell; tab names and order read back through AppleScript, cells of every sheet read back         | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean   | Opened from the shell; sheet names and order read back, sample cells read back                                   | 2026-11-04 |
| Google Sheets    | web      | browser  | pending | Needs a manual upload and a visual check                                                                         |            |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean   | Headless CSV conversion with the all-sheets token wrote one file per sheet, named after the sheet, each matching | 2026-11-04 |

## styles.xlsx (bold header and total row, a kroner currency column, a percent column, an ISO-date column, a default-formatted date column)

| Application      | Version  | Platform | Result  | How it was checked                                                                                                                                                                                  | Date       |
| ---------------- | -------- | -------- | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean   | Opened from the shell; every cell's value, number format and bold flag read back through AppleScript and matched the input                                                                          | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean   | Opened from the shell; formatted values read back: `1 234,50 kr`, `10,0 %`, `2026-11-01`, `11/3/26 14:30`                                                                                           | 2026-11-04 |
| Google Sheets    | web      | browser  | pending | Needs a manual upload and a visual check                                                                                                                                                            |            |
| LibreOffice Calc | 26.8.2.2 | macOS    | clean   | Headless HTML conversion shows `1,234.50 kr`, `10.0%`, ISO dates and bold header and total rows; its CSV exporter prints the currency column raw, a quirk of that filter, so format checks use HTML | 2026-11-04 |
