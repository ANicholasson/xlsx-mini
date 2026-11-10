# Auto column width, estimated from the cells and sampled when streaming

A column with `width: "auto"` takes the width of its widest cell; a sheet with `autoWidth: true` does that for every column without a width of its own. The width is an estimate from the text each cell shows, computed while the rows are rendered, with no font measurement and no dependency. On the byte path every row counts. When streaming, the first 1 000 rows count and the rest take the widths they set. A workbook that uses neither spelling renders byte for byte as before.

## Two spellings, one mechanism

`width: "auto"` on a `Column` is the form users of other writers look for and the only way to size one column among fixed ones. `autoWidth: true` on the sheet exists because the column list is not always known up front: `sheetFromObjects` without `columns` learns the keys from the first object, when the rows are already being read, and a plain `rows` sheet may be wider than its `columns` array. The sheet flag applies to every column that has no numeric width, so `columns: [{ width: 40 }]` with `autoWidth: true` fixes column A and sizes the rest. `sheetFromObjects` takes both: `autoWidth` in its options and `"auto"` on an object column. The error code `INVALID_COLUMN_WIDTH` keeps its meaning and now also names `"auto"` as the one accepted string.

## How a width is estimated

Excel stores a column width in characters of the default font's digit, Calibri 11 at 96 dpi, where a digit is 7 pixels and a column adds 5 pixels of padding. The estimate follows what Excel's own autofit does without opening a font: each printable ASCII character has a tabled pixel width in Calibri 11, other characters count 8 pixels, East Asian and other wide characters 14, combining marks nothing. A cell's pixels are the widest line of its text, scaled by its font size over 11 and by 1.1 when bold. The column's width is its widest cell plus 12 pixels, converted to characters and rounded to two decimals, no narrower than 3 and no wider than 60. The ceiling keeps one long note from making a column as wide as the screen, and a caller who wants more sets the number. A column whose cells are all empty writes no width and keeps the application's default.

What each kind of cell shows:

- **Text**, including a link's text and a boolean as `TRUE` or `FALSE`: the characters themselves.
- **A number** in the General format: its JavaScript text, at most 11 characters, which is what Excel shows before it shortens. With a number format: the length the format's first section produces, counted from its digit placeholders, its grouping comma, its decimals, its percent scaling, its exponent and its literal text, quoted or bare, with `_x` as a space and `*x` and `[...]` as nothing.
- **A date** or Temporal value: the length its format produces, with `mmmm` and `dddd` as nine letters, `mmm` and `ddd` as three, `yyyy` as four, every other token as two, and literals as themselves. The built-in formats 14 to 22 and 45 to 47 are tabled so an unstyled date measures as what it displays.
- **A formula**: its cached value, when given, as its own kind; a formula without one measures nothing, since the writer does not compute.
- **A merged cell** that spans more than one column measures nothing: its text lies across the merge, not in its first column. A merge down one column still counts.

Numbers and dates measure as digits at 7 pixels each; text goes through the table. No format engine is implied: the estimate reads the format string, it does not apply it. A wrapped cell measures like any other, so a long wrapped text widens its column up to the ceiling rather than staying at the default and wrapping into many lines; the ceiling is what bounds it.

## Where the measuring happens, and what it costs

The cell renderer already resolves a cell's value, type and merged style; the measurer is one call from there, made only for a column that is auto-sized and only while the sheet is still measuring. A text cell skips the table when its length times the widest tabled character cannot beat the column's current maximum, and a text that went through the shared-strings table is scanned once per sheet and found by its index after that, so a column of names or statuses costs a lookup per cell. A column stops measuring once it reaches the ceiling. A number cell that brings no style of its own is skipped when its magnitude does not exceed the largest the column has measured, since under one format the width grows with the magnitude. Nothing per cell is allocated: the measurer holds one slot per column with its pixel count, its resolved style factor and format shape, and the largest magnitude seen. On the wide-table suite with every column auto-sized, the STORE write of 100 000 rows costs the same as without, within noise, on Node and Bun, and fourteen percent more on Deno, against the ten percent the issue allowed; the many-sheets suite, where every one of forty sheets learns its widths from scratch, pays seven to thirteen percent, and the string-heavy suite nothing. The numbers are in ADR 0021. The three suites without auto width stay inside the noise band. The benchmark gains an `auto` variant next to `deflate`, `store` and `stream`: the same suite with every column auto-sized, written as STORE, so the cost of measuring is read against `store` in the same table. A base build that cannot write the variant is skipped on that row.

The byte path pays nothing for ordering: the worksheet head, which carries `<cols>`, is already written after the rows, so the widths are complete when it is rendered.

## What the streaming writer does

`<cols>` precedes `<sheetData>` in the schema, and the streaming writer has sent the head before the first row. Three ways out were open:

- Buffer the whole sheet when any column is auto-sized. Flat memory is the one promise streaming makes; a flag on a column that silently breaks it is the wrong kind of surprise.
- Refuse auto width when streaming, with a coded error. Honest, but it leaves the largest exports, the ones that stream, without the most-requested feature and with no way to get it: the caller cannot measure rows it has not read either.
- Measure a bounded sample of leading rows, hold the head and those rows, then send the head with the widths they set and the rows after it unchanged.

The third is what the writer does, with a sample of 1 000 rows counted from the first row of the sheet, header included. A sheet with fewer rows is measured in full and streams the same parts the byte path writes. For a longer sheet the first 1 000 rows set the widths and measuring stops, so the cost ends with the sample too. The held rows are rendered markup, not row objects, bounded by the sample and the row width, and they leave as one write when the head goes out; the first bytes of such a sheet are delayed until the sample has been read or the sheet has ended. This is the same approximation Excel makes on screen when it fits a column to the rows in view, and it is right whenever a column's content is of a kind, which is what tabular exports are. A column whose widest value sits past the first thousand rows shows it clipped, as a fixed width would; the README says so next to the option, and the sample size is a constant rather than an option until someone needs it to be one.

## What the change must also carry

A golden fixture `auto-width.xlsx` with text, numbers under formats, dates, booleans, a bold header, a long note that meets the ceiling and a merged title, opened in the four applications and recorded in the compatibility log with the check that no cell is clipped. The streamed test suite gains the fixture, which is shorter than the sample and so streams the same parts, and a test that a wider value past the sample does not change a width while one inside it does. A README section with a snippet run by the tests, the API table rows for `Column`, `Sheet` and `sheetFromObjects`, a sentence in the streaming section, a glossary entry, the benchmark variant and its number in ADR 0021. A minor changeset.

## Considered and rejected

- Measuring with a font: no dependency, no canvas, no file system in the core, and the applications differ in font anyway. The pixel table is what the writers that autofit use, and the ceiling and padding absorb the difference between Calibri, Helvetica and Arial.
- Character count times a constant, the simplest estimate: `Wim` and `illi` are the same five characters and twice as wide; headers in capitals would clip.
- A `widths: "auto"` option on `writeXlsx` for the whole workbook: the sheet is where the columns are described, and the flag reads next to `autoFilter`.
- Measuring the full streamed sheet and sending `<cols>` after the rows: the schema fixes the order and Excel repairs a file that breaks it.
- A larger or configurable sample: a thousand rows cost about a megabyte of held markup on the wide table and a fraction of a second from a cursor; nobody has asked for a different number, and a constant can become an option without a breaking change.
- Per-format rendering of numbers for an exact length: a number-format engine for the sake of a width estimate, when the placeholders already say how many characters come out.
