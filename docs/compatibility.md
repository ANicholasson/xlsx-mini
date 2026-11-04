# Compatibility log

Which applications opened each golden fixture without a repair prompt, and how that was checked.
A row is only "clean" when the application opened the file without any dialog and the cell values
read back as written. Add a row per application and fixture whenever a fixture changes.

## minimal.xlsx (one sheet, strings and numbers)

| Application      | Version  | Platform | Result  | How it was checked                                                                                                                                         | Date       |
| ---------------- | -------- | -------- | ------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------- |
| Microsoft Excel  | 16.114.3 | macOS    | clean   | Opened from the shell; all 15 cells read back through AppleScript (`value of range "A1:C5"`), including escaped text, emoji, padded whitespace and `1e+21` | 2026-11-04 |
| Apple Numbers    | 14.4     | macOS    | clean   | Opened from the shell; imported table read back through AppleScript cell by cell                                                                           | 2026-11-04 |
| Google Sheets    | web      | browser  | pending | Needs a manual upload to Drive and a visual check                                                                                                          |            |
| LibreOffice Calc |          |          | pending | Not installed on the development machine yet; the CI validate job will convert fixtures to CSV headlessly                                                  |            |

Independent zip checks on the same fixture: `unzip -t` (Info-ZIP 6.00) and Python's `zipfile` both report no errors.
