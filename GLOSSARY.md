# xlsx-mini

Writes workbooks as `.xlsx` files. One context: the spreadsheet as the caller describes it, and the package Excel reads.

## Language

### What the caller describes

**Workbook**:
The whole document handed to the writer: an ordered list of sheets.
_Avoid_: file, document, spreadsheet

**Sheet**:
One tab of a workbook, with a name and rows.
_Avoid_: worksheet (reserved for the package part), tab, table

**Row**:
One horizontal line of cells, left to right from column A.

**Cell**:
One position in a row, holding a value and optionally a style.
_Avoid_: field

**Value**:
What a cell holds: text, a number, a boolean, a date-like value, or nothing.

**Empty cell**:
A position with nothing written; it keeps its column so the cells after it stay in place.
_Avoid_: blank, null cell

**Column**:
A vertical position across every row, named by letters, carrying a width and a default style for its cells.

**Style**:
How a cell looks: font (bold, italic, size, colour), fill, border, alignment, wrap, indent and number format. A cell's own style wins over its column's, property by property.
_Avoid_: formatting, format (reserved for number format)

**Number format**:
Excel's display rule for a numeric value, such as a currency, a percentage or a date layout; a built-in id or a format string.
_Avoid_: numFmt, formatting

**Merge**:
A range of cells shown as one; the top-left cell's value and style show.
_Avoid_: span, merged region

**Row height**:
The height of a row in points, set by row number; rows without one take Excel's default.

**Streamed workbook**:
A workbook as `streamXlsx` takes it: the same sheets, whose rows may arrive asynchronously, so a cursor can feed a sheet page by page.
_Avoid_: lazy workbook, async workbook

**Formula**:
A cell that holds an expression the application computes when the file opens, written without the leading `=`; an optional cached value shows until then. Never inferred from text.
_Avoid_: function, calculation

**Hyperlink**:
A cell whose text opens an external URL or jumps to a cell in the workbook. External targets live in the worksheet's relationships part.
_Avoid_: link target, URL cell

**Data validation**:
A rule on the cells of a range: a dropdown from a list or a range, or a number or date within bounds, with its input and error messages.
_Avoid_: dropdown (one kind of rule), constraint

**Print setup**:
How a sheet prints: orientation, paper size, fit to pages, margins, header and footer, repeated rows.
_Avoid_: page layout, page setup (reserved for the element)

**Freeze panes**:
The rows and columns that stay visible while the sheet scrolls.

**Autofilter**:
Filter dropdowns on the header row across the used range.

### What Excel reads

**Package**:
The zip container that an `.xlsx` file is.
_Avoid_: archive, zip

**Part**:
One file inside the package: workbook, worksheet, styles, shared strings, relationships, content types.

**Shared string**:
Text stored once in a workbook-wide table and referenced by index from every cell that uses it. What `writeXlsx` writes.

**Inline string**:
Text written inside its own cell, so nothing has to be held until the end. What `streamXlsx` writes unless asked for shared strings.
_Avoid_: rich string

**Chunk**:
One piece of a streamed part as the consumer receives it: at most 65 535 bytes, so it is one stored deflate block where that fallback applies.

**Data descriptor**:
The CRC and sizes of a streamed entry, written after its data because they are only known then; readers find them again in the central directory.

**Serial**:
Excel's number for a date and time: days since 1899-12-30, with the time of day as the fraction.
_Avoid_: timestamp, epoch days

**Repair dialog**:
Excel's "We found a problem with some content" prompt. The failure this project exists to never cause.
_Avoid_: corruption warning

### How the project checks itself

**Golden fixture**:
A committed `.xlsx` file the writer must reproduce byte for byte from a known input.
_Avoid_: snapshot

**Compatibility log**:
The record of which applications opened each golden fixture clean, and how that was checked.

**Runtime matrix**:
The hosts the built package must write a workbook on at every test run: Node (ESM and CommonJS), Bun, Deno, Cloudflare workerd and the Vercel edge runtime. STORE output is byte-identical across them.

**Benchmark suite**:
One of the three workbooks in `bench/` the writer is timed on: the wide table, the many-sheets workbook and the string-heavy workbook. Each is generated from a fixed seed.

**Budget**:
The time and peak-memory ceiling for a benchmark suite, recorded in the baseline ADR. A budget only ratchets down.

**Benchmark report**:
The one comment on a pull request with a table per runtime of the pull request against `main`, both timed on the same runner in turns. Updated on every push.
