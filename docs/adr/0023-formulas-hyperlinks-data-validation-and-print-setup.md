# Formulas, hyperlinks, data validation and print setup as explicit, additive elements

Four features that each map to one element of the worksheet part, declared where the caller already describes the sheet: formulas and hyperlinks on the cell, validations and print setup on the sheet. A workbook that uses none of them renders byte for byte as before, on both the byte path and the streamed path.

## Formulas are a cell type, never inferred

A formula cell is `{ formula: "SUM(B2:B4)" }`, optionally with a cached `value` and a `style`. Text that starts with `=` stays text, as ADR 0019 promised, so pasted user data can never become a formula. The formula is written as given inside `<f>`, escaped for XML and never evaluated: the writer has no formula engine and no view of other sheets' data. A formula that starts with `=` is an error rather than being stripped, since Excel repairs a file whose `<f>` carries one; an empty formula or one over Excel's 8 192 characters is an error too. The cached value is written as the cell's value would be, typed `str` for text since a formula result is never a shared string, so the two paths and both string modes agree.

When any sheet has written a formula, the workbook part carries `<calcPr fullCalcOnLoad="1"/>`: Excel then computes every formula on open instead of showing empty cells, which is what the other applications do anyway. The flag makes Excel ask to save on close, as every writer without a calculation engine accepts; a workbook without formulas has no `calcPr` and no prompt.

## Hyperlinks live in the sheet, external targets in its relationships

An external link is `{ value, href }`; `href` must parse as an absolute URL, so a relative path or a bare host is an error and the file never carries a target Excel would refuse. The target is a relationship in `xl/worksheets/_rels/sheetN.xml.rels`, a part that exists only for a sheet with external links and is written straight after that sheet on both paths, so the streamed package needs no second pass. An in-workbook link is `{ value, cell: "B2" }` within the same sheet or `{ value, sheet: "Summary", cell: "B2" }` to another, written as a `location` without a relationship; the sheet must exist in the workbook, compared without case as Excel does, and the cell is a cell or range in A1 form. The link's text is the cell's value, escaped and interned or inlined like any other.

No link look is applied: a style the caller did not ask for and cannot remove is not declarative, and Excel itself shows a link written this way in the cell's own style. `underline` joins `CellStyle` so a link can look like one.

## Data validation and print setup are sheet-level declarations

`validations` is a list of rules on the sheet, each with a `range`: a cell, a range in A1 form, or whole columns such as `B:B`, written as rows 1 to 1 048 576 since that is how Excel stores a whole column. A rule is one of four kinds: a dropdown from a literal list (`list: ["Open", "Closed"]`, within Excel's 255 characters for the joined source), a dropdown from a range (`list: "Lists!A2:A9"`, a range on this or a named sheet), a number range (`number: { min, max, integer }`), or a date range (`date: { min, max }`). Each takes an optional input message and error message, with Excel's limits of 32 characters for a title and 255 for a message, and `allowBlank`, true by default as in Excel's dialog. Rules are validated up front and written after the rows, so a sheet whose rows come from a generator needs nothing held but the rules themselves.

`print` carries orientation, a paper size by name, fit to a number of pages wide or tall, margins in inches as Excel stores them, header and footer text with Excel's own codes (`&P`, `&N`, `&D`, `&L`, `&C`, `&R`, `&&`) and `repeatRows`, the header rows repeated on every page. The first three become `pageSetup`, with `fitToPage` in `sheetPr` when fitting; margins are always written with Excel's defaults filled in once `print` is set, since Excel writes them before every `pageSetup`; repeated rows become the `_xlnm.Print_Titles` defined name in the workbook part, next to the autofilter's name. The text is passed through as Excel's syntax rather than wrapped in a structured object, because the codes are the feature and the syntax is documented in every application.

## One range parser, one error code per feature

Cell and range references are parsed by one function in the worksheet slice, shared by merges, hyperlinks and validations, so "not A1 form" is reported the same way everywhere. Each feature has its own code (`INVALID_FORMULA`, `INVALID_HYPERLINK`, `INVALID_VALIDATION`, `INVALID_PRINT_SETUP`) and the path names the cell, rule or property, as ADR 0012 requires.

## What the change must also carry

A golden fixture per feature, opened in the four applications and recorded in the compatibility log, with LibreOffice reading the formula results back in CI. README sections with tested snippets, the API and error tables, the glossary. The three benchmark suites inside the noise band, since the cell object check on the byte path grows by a property test. A minor changeset.

## Considered and rejected

- Inferring a formula from text starting with `=`: ruled out in ADR 0019; data must never become code.
- Stripping a leading `=` instead of refusing it: a silent rewrite of the caller's input, where every other bad input throws.
- Array and shared formulas, and defined names: not asked for; each is additive later.
- Computing formula results for the cache: a formula engine is a different library.
- A default hyperlink style, or Excel's built-in `Hyperlink` cell style: a look the caller did not declare; `underline` lets them declare it.
- Accepting any string as `href`: a relative target or a bare host writes a relationship Excel repairs.
- Validation by custom formula or text length, and print area, page breaks, scaling and gridlines: not asked for; additive later.
- A structured header and footer (`{ left, center, right }`): would still need the codes for page numbers and dates, so it would be a second syntax on top of Excel's.
- Validations in `Column`: a column's rule usually starts below the header, which a range says and a column slot cannot.
