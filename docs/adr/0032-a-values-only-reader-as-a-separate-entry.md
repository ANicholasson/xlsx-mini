# A values-only reader as a separate entry

The package gains a second direction: `xlsx-mini/read` reads the values of an `.xlsx` file. It reads values and nothing else, it is built for files from strangers, and it is experimental while the version is 0.x. Its presence costs an application that imports only the writer nothing: erase the read side from the build and the writer's bundles are the same bytes.

```ts
import { XlsxReadError, objectsFromRows, openXlsx } from "xlsx-mini/read";

const book = await openXlsx(file); // Uint8Array, ArrayBuffer or Blob
book.sheets; // ["Orders", "Customers"]
for await (const row of book.rows("Orders")) console.log(row.number, row.cells);

for await (const { number, values } of objectsFromRows(book.rows(0), columns)) {
  if (typeof values.total !== "number") console.log(`Row ${number}: Total must be a number`);
}
```

This reopens the first non-goal of ADR 0026 by the rule that ADR set: a design that leaves the bundle of an application that does not use the feature unchanged, and that can be proven. Editing an existing workbook, charts, pivot tables and password encryption stay non-goals. ADR 0001's "write-only" now describes the entry `xlsx-mini`, not the package.

## Why

An application that exports a sheet for a person to fill in needs to take it back. Without a reader in this package it keeps a second spreadsheet library for the way in, which is the dependency this package exists to remove. A survey in November 2026 found no small reader built on web-standard APIs with no dependencies: the readers measured were between 18 and 42 kB gzipped, or two weeks old.

Three throwaway prototypes were built against one contract and scored by one harness before this decision. Each read 84 workbooks written by this library, Excel, Numbers and LibreOffice, refused or survived 112 hostile files and 6 000 mutated ones, and came to between 5.0 and 6.0 kB gzipped. All three chose the same structure without being told. The numbers below come from them.

## What it gives

- `openXlsx(file, { limits })` reads the package's directory and the sheet list. `sheets` holds every sheet's name in workbook order, hidden ones included.
- `rows(sheet)` takes an exact name or an index from 0 and gives each row that holds a value, in order, as `{ number, cells }`: the sheet's row number from 1, and the values from column A with `null` for an empty cell and no `null` at the end. Nothing of a sheet is read until its rows are asked for, and memory stays flat however long the sheet is.
- `objectsFromRows(rows, columns)` is the mirror of `sheetFromObjects`. The first row is the header; the caller names each key and its header text; a header the file lacks is an error. No object is ever keyed by text from the file. The column list has the shape of the writer's, so one list serves the export and the import. Saying what a value must be stays with the caller.
- Input is a `Uint8Array`, an `ArrayBuffer` or a `Blob`. A `Blob` is read in slices and never whole. There is no stream input: a package's directory is at its end.

A value is a string, a number, a boolean, a `Date` or `null`. Text is as stored, untrimmed, with rich text as its plain text; empty text is `null`. An error cell is its text. A formula is its stored result, or `null` when the file stores none; nothing is calculated. Hidden sheets, rows and columns read like visible ones, and a merge has its value in its top-left cell. Pictures, charts, comments and every other part the reader has no use for are never inflated, so a workbook that holds them opens like any other.

## Dates

A number is a date when the number format of its cell style shows a year, a month or a day. The format code is tokenized, with no regular expression; a custom code replaces a built-in id, 0 included. Only the first section counts: a serial that can be a date is positive, and a positive number shows in the first section. `y` and `d` are date tokens in either case, and so is `mmm` or longer, a month's name; `m` and `mm` are a month unless an hour or a second comes before them or a second after them. Quoted text, an escaped character, the character after `_` or `*` and a bracketed section hold no token, and a `[` inside a bracketed section ends what is read, as LibreOffice shows `[[h]]:mm` as a number and `yyyy[[h]]` as a date. `General` and the exponent of `0.00E+00` are not dates: LibreOffice and Numbers write `General` as a custom code, and reading its letters as era tokens turned 28 of 84 real workbooks' numbers into dates in a trial. So no other letter is a date token, the era letters `e` and `g` and the Buddhist year `b` included: LibreOffice shows a code of `e` or `g` alone as a date and openpyxl as a number, and an era format such as `ge.m.d` is a date by its `m` and `d`. A time of day or an elapsed time stays a number, a fraction of a day. Of the built-in ids, 14 to 17, 22, 27 to 31, 36 and 50 to 58 are dates. Ids 34, 35 and 52 to 56 mean a date in some East Asian locales and a time in others, and the file does not say which wrote it; reading 52 to 56 as dates and 34 and 35 as times is a guess that is recorded here as one. The Thai ids 59 to 81 are not dates unless the file defines them so, though LibreOffice reads 71 to 75 and 78 as dates.

Only a number becomes a date: text, a boolean and an error keep their type under a date style, and a formula's stored number is a date. The styles part is read once per opened workbook, by the first read of a worksheet, and kept as one byte per cell style, or not kept when no cell style shows a date. A cell's style is read only where it can change the value: in a cell that holds a number, in a workbook where some cell style shows a date. There a style that does not exist is refused. Anywhere else the style is not read and so not refused, the rule the reader follows for the parts it never opens. A styles part that LibreOffice and openpyxl read in different ways is refused: two lists of cell styles, a list of formats or of cell styles below an `AlternateContent`. A second list of formats, one after the cell styles, an id written other than in decimal digits alone, and anything in either list that is not one of its entries are refused too. A style that names an id the list of formats does not define, past the built-in ones that end at 81, shows no date, so its cells read as the numbers they store: openpyxl reads ids 82 to 163 as `General`, and LibreOffice showed a number in three of four such parts made by hand, and a date in the fourth, which also had a list of formats and gave the id a date code in a differential format after the cell styles.

A date comes back as a `Date` whose UTC fields equal the sheet's wall clock, rounded to the millisecond, which is what the writer does in the other direction (ADR 0009). A `Date` is given exactly for the serials the writer can write. A serial below 1, the serial of 29 February 1900, a negative serial and one past the year 9999 stay numbers: Excel shows the first as day 0 of January and the second as a day that never was, so no calendar date is true to them. The range is checked after rounding to the millisecond, so a serial less than half a millisecond below 60 or below the end of 9999 stays a number too. A workbook in the 1904 date system says so with `date1904` of its properties, `1` or `true`; `0`, `false` and no attribute mean 1900, and any other spelling is refused, since LibreOffice and openpyxl read `TRUE`, `yes` and `off` in different ways. There, serial 0 is 1904-01-01, a negative serial stays a number, and no day is missing. Every `Date` read is one the writer writes back.

A cell of type `d` holds its date as text, and the type makes it a date whatever its style (ECMA-376 part 1, 18.18.11). The forms read are XML Schema's `date` and `dateTime` with a year of four digits, any number of decimals of a second, `Z` or no time zone, and white space around them, which XML Schema collapses, as in `2026-11-17` and `2026-11-17T12:34:56.5`. Decimals past the third round the time to the millisecond, half up, as the writer rounds a serial, and the range is checked after the rounding. Anything else is refused as `INVALID_CELL`: an offset, which LibreOffice converts and openpyxl drops, a day that does not exist, a year before 1900, and a time that rounds up into the year 10000.

## Refuse, never repair

The reader is handed files by strangers. Other readers have been hurt by prototype pollution through names in the file, by patterns that backtrack, by a damaged directory that loops, and by one cell reference in the far corner. So:

- Anything malformed is refused with a code: rows or cells out of order, a bad reference, a number that is not decimal, text that is not valid UTF-8, a DOCTYPE, two parts with one name, a local header that disagrees with the directory, a part that is not the size or the CRC-32 its directory entry states.
- A workbook whose sheet list or sheet data cannot be found is an error, never an empty result. In the prototypes each reader read a few damaged workbooks as empty, which for an import is the worst outcome: nothing fails and nothing arrives. A sheet without cells by its kind (a chart sheet, a dialog sheet, a macro sheet) is listed and has no rows; a sheet whose relationship is of any other type than those and a worksheet is refused, because one damaged character in that type otherwise reads as an empty sheet.
- Two sheets with the same name, character for character, are refused: a caller chooses a sheet by its exact name. Names that differ only by case or by Unicode folding are two sheets, as LibreOffice writes them.
- Nothing is sized from what the file claims, no name from the file becomes a property, and nothing outside the package is resolved or fetched.
- Zip64 packages are accepted.

## Limits

Every limit has a default sized for a file a person uploads. A caller can raise one per `openXlsx` call; no value turns one off. The byte, position and row limits are per read: two sheets read at the same time each have their own, so what a caller holds in memory is the sum of the reads it runs. One read holds one row and the text of its part around it. In trials with the default limits, the most one read reached was 106 MiB of heap on Node, for a row of six cells that each hold the longest text of character references, and 309 MiB resident on Bun, for 30 rows that each hold one cell of 4 194 304 carriage returns. These are measurements, not bounds, and they move with when the engine collects: that row reached 88 to 106 MiB over 19 runs. The file decides how many sheets there are, and any number of them may name one part: a file of 6 KB whose 150 sheets name one part, a row of 4 194 304 characters, exhausted a 512 MiB heap when every sheet was read at once. No limit spans reads. A caller that reads the sheets of a file from a stranger one at a time holds one read's worth at a time, and beside it the shared-string table, which the opened workbook holds for as long as the caller keeps the workbook.

| Limit               | Default   | What it bounds                                                                                                                                                       |
| ------------------- | --------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `inflatedBytes`     | 128 MiB   | What one operation inflates. Opening the workbook, with its shared strings and styles, is one operation; each `rows` iteration is one.                               |
| `cellPositions`     | 5 000 000 | The positions one `rows` iteration yields, empty cells before a value included.                                                                                      |
| `rowCharacters`     | 4 194 304 | The text one row holds. A row is held whole, so this bounds the text of the values one read holds at a time.                                                         |
| `sharedStrings`     | 1 000 000 | Entries in the shared-string table.                                                                                                                                  |
| `sharedStringBytes` | 32 MiB    | What the shared-string table inflates to. Held, in trials, at up to 2.9 times this on Node and 4.2 on Bun (a million entries of 16 characters, one outside Latin-1). |
| `parts`             | 100 000   | The parts in the package, read or not. A workbook with one picture per row has about one part per row.                                                               |

A part's declared size is charged before a byte is inflated, then enforced exactly. Parts that become lists in memory are held to an eighth of `inflatedBytes`, and so is the package's directory, whose size is charged before it is read. Two more rules are fixed and are not limits: elements nest at most 256 deep and one token is at most 4 194 304 characters, reported as `INVALID_XML`. No workbook comes near either, and without them a file inside every limit above exhausted a 512 MiB heap. The row and parts limits exist for the same reason: three sheets of one 119 MiB row each, and a directory of four million entries, did the same in trials.

Rows are yielded before a sheet's size and checksum can be known, at its last byte. A caller acts on what it read after the loop has ended without an error.

## Inflate comes from the platform

Parts are inflated by `DecompressionStream`, as the writer deflates by `CompressionStream` (ADR 0003). Where the platform has none, a deflated part is refused with `NO_INFLATE` and stored parts still read. An inflate of its own would be the largest part of the reader and the part most likely to be wrong.

## Where it lives, and what it shares

The read side of a feature is its root file `read.ts` and a private folder `read/`; the package entry is `src/read.ts`. So the package reader is beside the package writer in `opc/`, the tokenizer beside the escaping in `xml/`, and each part's reader beside its writer. This is ADR 0013 applied, not amended: a feature owns a piece of the format in both directions.

The two directions are one library. What both need is written once and imported by both: the CRC-32, the zip records, the serial arithmetic, the bounds of a sheet. A shared fact gets its one home in a change to the writer of its own, with its byte delta stated, before the reader needs it. A change that builds the reader does not edit a file the writer bundles.

Four checks hold the promise to the writer's users. A boundary rule forbids anything outside the read side from importing it, tests and benchmarks excepted. The size test bundles the writer's entries, erases the read side from a copy of the build, bundles again and requires the same bytes. A digest of each writer bundle is committed, and the test fails when a build's digest differs: the first two checks compare a build with itself, so only this one sees a change made through a shared file, and it makes every change to the writer's bytes a line in the pull request that causes it. The reader has a budget of its own, 12 000 bytes gzipped, beside the writer's two (ADR 0021).

## Errors

The reader throws `XlsxReadError`, never `XlsxMiniError`. It has `code` from its own union and, where known, `sheet`, `cell`, `limit` and `header`. The writer's error carries a path into the caller's input (ADR 0012); a reader has no such input, and what its caller needs is where in the file. The project now has two error classes, one per direction.

## Experimental

`xlsx-mini/read` may change in any 0.x minor. Every export says so in its documentation and the README section says so in its heading; nothing is said at run time. The bar for 1.0 (ADR 0020) includes the reader no longer being experimental.

## How it is proven

- Round trip: every golden fixture is read back and compared with its input.
- Differential: a committed corpus written by Excel, Numbers, LibreOffice and Google Sheets is read and compared with LibreOffice's typed read-back of the same files, with the known places where that read-back is not the truth written down.
- A list of hostile files, each with the outcome it must have, and seeded fuzzing, both under time and memory bounds.
- The runtime matrix reads back what each host wrote.

## Considered and rejected

- A package of its own: a second name for one format, and a second copy of everything the two share.
- The read side as feature folders of its own: a simpler wall, paid for with every format fact stated twice in features that never import each other.
- Reading styles, formulas, comments or pictures: each is a reader of its own, and a value is what an import needs.
- An object keyed by the header row: keys from the file are how prototype pollution starts, and a typed object needs declared keys anyway.
- `book.objects(sheet, columns)` as a method: a caller who reads rows only would pay for it, and the function mirrors `sheetFromObjects` name for name.
- An inflate of its own, for the platforms without one.
- Repairing: following a best guess past a malformed row gives a caller values no application would show.
- Limits that can be switched off.
