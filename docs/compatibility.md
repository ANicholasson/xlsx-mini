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
