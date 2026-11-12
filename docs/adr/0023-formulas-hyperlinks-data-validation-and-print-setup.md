# Formulas, hyperlinks, data validation and print setup as explicit, additive elements

Four features that each map to one element of the worksheet part, declared where the caller already describes the sheet: formulas and hyperlinks on the cell, validations and print setup on the sheet. A workbook that uses none of them renders byte for byte as before, on both the byte path and the streamed path.

## Formulas are a cell type, never inferred

A formula cell is `{ formula: "SUM(B2:B4)" }`, optionally with a cached `value` and a `style`. Text that starts with `=` stays text, as ADR 0019 promised, so pasted user data can never become a formula. The formula is written as given inside `<f>`, escaped for XML and never evaluated: the writer has no formula engine and no view of other sheets' data. A formula that starts with `=` is an error rather than being stripped, since Excel repairs a file whose `<f>` carries one; an empty formula or one over Excel's 8 192 characters is an error too. The cached value is written as the cell's value would be, typed `str` for text since a formula result is never a shared string, so the two paths and both string modes agree.

When any sheet has written a formula, the workbook part carries `<calcPr fullCalcOnLoad="1"/>`: Excel then computes every formula on open instead of showing empty cells, which is what the other applications do anyway. The flag makes Excel ask to save on close, as every writer without a calculation engine accepts; a workbook without formulas has no `calcPr` and no prompt.

## Hyperlinks live in the sheet, external targets in its relationships

An external link is `{ value, href }`; `href` must parse as an absolute URL, so a relative path or a bare host is an error and the file never carries a target Excel would refuse. The schemes that run code rather than open something (`javascript:`, `vbscript:`, `data:`) are refused too; every other scheme is allowed, since `file:`, `tel:` and an application's own scheme are ordinary link targets in a workbook. The target is a relationship in `xl/worksheets/_rels/sheetN.xml.rels`, a part that exists only for a sheet with external links and is written straight after that sheet on both paths, so the streamed package needs no second pass. An in-workbook link is `{ value, cell: "B2" }` within the same sheet or `{ value, sheet: "Summary", cell: "B2" }` to another, written as a `location` without a relationship; the sheet must exist in the workbook, compared without case as Excel does, and the cell is a cell or range in A1 form. The link's text is the cell's value, escaped and interned or inlined like any other.

No link look is applied: a style the caller did not ask for and cannot remove is not declarative, and Excel itself shows a link written this way in the cell's own style. `underline` joins `CellStyle` so a link can look like one.

## Data validation and print setup are sheet-level declarations

`validations` is a list of rules on the sheet, each with a `range`: a cell, a range in A1 form, or whole columns such as `B:B`, written as rows 1 to 1 048 576 since that is how Excel stores a whole column. A rule is one of four kinds: a dropdown from a literal list (`list: ["Open", "Closed"]`, within Excel's 255 characters for the joined source), a dropdown from a range (`list: "Lists!A2:A9"`, a range on this or a named sheet), a number range (`number: { min, max, integer }`), or a date range (`date: { min, max }`). Each takes an optional input message and error message, with Excel's limits of 32 characters for a title and 255 for a message, and `allowBlank`, true by default as in Excel's dialog. Rules are validated up front and written after the rows, so a sheet whose rows come from a generator needs nothing held but the rules themselves.

`print` carries orientation, a paper size by name, fit to a number of pages wide or tall, margins in inches as Excel stores them, header and footer text with Excel's own codes (`&P`, `&N`, `&D`, `&L`, `&C`, `&R`, `&&`) and `repeatRows`, the header rows repeated on every page. The first three become `pageSetup`, with `fitToPage` in `sheetPr` when fitting; margins are always written with Excel's defaults filled in once `print` is set, since Excel writes them before every `pageSetup`; repeated rows become the `_xlnm.Print_Titles` defined name in the workbook part, next to the autofilter's name. The text is passed through as Excel's syntax rather than wrapped in a structured object, because the codes are the feature and the syntax is documented in every application.

## One range parser, one error code per feature

Cell and range references are parsed by one function in the worksheet slice, shared by merges, hyperlinks and validations, so "not A1 form" is reported the same way everywhere. Each feature has its own code (`INVALID_FORMULA`, `INVALID_HYPERLINK`, `INVALID_VALIDATION`, `INVALID_PRINT_SETUP`) and the path names the cell, rule or property, as ADR 0012 requires.

## What the change must also carry

A golden fixture per feature, opened in the four applications and recorded in the compatibility log, with LibreOffice reading the formula results back in CI. README sections with tested snippets, the API and error tables, the glossary. The three benchmark suites inside the noise band, since the cell object check on the byte path grows by a property test. A minor changeset.

## Functions newer than Excel 2007, 2026-11-12

A formula was stored exactly as given. Functions Excel added after 2007 are stored in a file under a prefixed name, `_xlfn.XLOOKUP`, and for `FILTER` and `SORT` `_xlfn._xlws.FILTER`; the applications show the bare name. A workbook written by 0.6.1 with such functions as a caller types them, and the same formulas prefixed by hand, read back as follows:

| Formula                               | Excel 16.114 (macOS)      | LibreOffice 26.8 | Numbers 14.4 |
| ------------------------------------- | ------------------------- | ---------------- | ------------ |
| `STDEV.S`, as typed                   | `#NAME?`                  | `#NAME?`         | empty        |
| `IFS`, `TEXTJOIN`, `CONCAT`, as typed | `#NAME?`                  | `#NAME?`         | computed     |
| `XLOOKUP`, `MAXIFS`, as typed         | `#NAME?`                  | `#NAME?`         | computed     |
| `SUM(FILTER(…))`, as typed            | empty                     | `#NAME?`         | computed     |
| `COUNTA(UNIQUE(…))`, as typed         | 1, the count of one error | `Err:509`        | computed     |
| Every one of them, prefixed           | computed                  | computed         | computed     |

So a formula with any of these was wrong in two of the three applications until its cell was edited, without a repair dialog and without a word. The writer now stores each call of such a function under its prefixed name, so a formula works as it is typed. Three further readings in Excel set the limits:

- `_xlfn.LET(x,A2,x*2)`, the function prefixed and its parameter not, is a file Excel repairs: `LET` and `LAMBDA` name their own parameters, and each is stored as `_xlpm.x`. Telling a parameter from a name takes a formula parser, which the writer does not have. A formula that calls `LET` or `LAMBDA` is therefore written exactly as given, every other call in it included, so the writer never prefixes by halves; the caller writes those two in their stored form.
- A function that returns an array, `UNIQUE`, `SEQUENCE`, `FILTER`, as a cell's whole formula shows its first value and fills no cell below. Spilling needs the cell marked as an array formula with metadata beside it. That is not written: such a function works inside one that takes an array, as `SUM(FILTER(…))` does.
- A name already prefixed, in either case, is left alone, and so is anything inside double quotes or a sheet name in apostrophes.
- The limit of 8 192 characters is on the formula as stored: Excel 16.115 opens a file whose stored formula has 8 192 and repairs one with 8 193, where the formula as typed had 5 307. The writer therefore measures a formula with its prefixes, and refuses one that the prefixes carry over the limit.

Google Sheets, given the golden fixture, drops `_xlfn.` and computes `STDEV.S`, `IFS`, `TEXTJOIN`, `XLOOKUP` and `MAXIFS`. Two rows differ there:

- `SUM(FILTER(…))` shows `#NAME?`, "Unknown function: '_xlws.FILTER'". Stored as `_xlfn.FILTER`, without the `_xlws.`, the formula computes in Excel 16.115 and shows `#NAME?` in LibreOffice 26.8, and Excel saving that workbook writes `_xlfn._xlws.FILTER` back. No stored name works in all four applications, so the writer stores Excel's, and Sheets misses `FILTER` here as it would in a file Excel saved. `SORT` is stored under the same prefix and was not checked there.
- `CONCAT(E2:E4)` shows `#N/A`. The name is read; Sheets' own `CONCAT` takes two values and not a range. That is a difference between the two applications' functions, which no stored name changes.

The names are Excel's own list of functions stored this way, 150 of them once `LET`, `LAMBDA`, the two operator stand-ins and the Python functions are left out, as Excel's documentation and the two writers that follow it publish them. A function added to Excel later needs adding here; until then it behaves as every one of them did before. The scan is one pass over a formula that contains a parenthesis, on formula cells only. The list costs the bundle 1 136 bytes gzipped, seven percent, for every application whether it writes formulas or not. A helper the caller wraps a formula in would have kept that out of other bundles, and would have left the default wrong; a formula that shows `#NAME?` only in some applications is the kind of failure this project exists to prevent, so the default pays.

## Considered and rejected

- Inferring a formula from text starting with `=`: ruled out in ADR 0019; data must never become code.
- Stripping a leading `=` instead of refusing it: a silent rewrite of the caller's input, where every other bad input throws.
- Array and shared formulas, and defined names: not asked for; each is additive later.
- Computing formula results for the cache: a formula engine is a different library.
- A default hyperlink style, or Excel's built-in `Hyperlink` cell style: a look the caller did not declare; `underline` lets them declare it.
- Accepting any string as `href`: a relative target or a bare host writes a relationship Excel repairs.
- A branded or constructed formula cell (`formula("SUM(A1)")`, a class) so that no plain object can be one: the plain-object cell is the API's shape since ADR 0005, and the injection the project guards against is text, which never becomes a formula; a caller who puts untrusted objects into cells must convert them to text, which the README says.
- Validation by custom formula or text length, and print area, page breaks, scaling and gridlines: not asked for; additive later.
- A structured header and footer (`{ left, center, right }`): would still need the codes for page numbers and dates, so it would be a second syntax on top of Excel's.
- Validations in `Column`: a column's rule usually starts below the header, which a range says and a column slot cannot.
