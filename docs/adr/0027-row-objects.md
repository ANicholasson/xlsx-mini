# Row objects: a style and a height on the row

A row is an array of cells, or an object `{ cells, style?, height? }`. The object is for what belongs to the row itself: a header, a band, a total row take one style for the row where they took one per cell. Keeping the look apart from the data is the request with the most support in the trackers of other spreadsheet writers, and a row was the one place this library had no slot for it: a style belonged to a cell or to a column.

## The shape

`Row` is `readonly Cell[] | RowObject`. An array stays the ordinary row and costs what it cost. The object names its cells, so nothing about a row is guessed from its content, and it is the home for what a row carries from here on: its style and its height now, whether it is hidden next.

A sheet-level `rowStyles` by row number, beside `rowHeights`, was the alternative. A row from a generator or a database cursor does not know its number, and the rows that need a style are picked by what they hold, not by where they land, so the style travels with the row.

## Where the style reaches

Every position in `cells`, the empty ones included, and nothing past the last. `{ cells: ["Total", null, 15196], style }` writes three styled cells, the middle one empty, so a fill or a border runs across the blank; a fourth column is untouched. The caller sets how far a band reaches by how long the row is.

Excel's own row format, `<row s="…" customFormat="1">`, reaches to the last column of the sheet. A header band that runs past its table into sixteen thousand empty columns is rarely what a table wants, and what the other applications do with a row format on cells that are absent is theirs to decide. Cells that are written are read the same everywhere: the golden fixture opens alike in the four applications.

A bare `null` in an array row stays nothing, also in a styled column, as before. The row style differs on purpose: a column runs down a million rows and a row ends where its cells end.

## Precedence

A cell's own style wins over its row's, and its row's over its column's, property by property, as a cell's won over its column's. A bold total row keeps the number format of each column; a cell in a band can take a fill of its own. A property set to `undefined` counts as not set.

The style table merges two styles, a cell's and a column's, and caches the result by the identity of the two objects. A third level did not need a third argument: for a styled row, the row's style merged over each column's stands in for the column styles. That merge happens once per distinct row style object and column, kept per sheet in a `WeakMap`, and the cells of the row then take the path every cell takes. A row style is validated when it is first met, so a wrong one is reported at the row, `sheets[0].rows[3].style.fill`.

So a row given as an array pays one `Array.isArray`, and the loop it runs is the one it ran. A row object runs a loop of its own, which is why the two are apart in the code. As with a cell style, a style object made anew for every row is merged anew for every row: rows that share a style should share the object.

## Height

`height` is in points, from above 0 to 409 as in `rowHeights`, and wins over `rowHeights` for the same row, as the more particular of the two; both win over the height a picture gives its row. `rowHeights` stays: it sizes a row by number without touching the rows, and reaches rows past the last one given.

## From objects

`sheetFromObjects` takes `rowStyle: (item, index) => CellStyle | undefined`, the index counting the objects from 0. A band is `index % 2`, and a row that needs attention is a condition on its object. An object without a style stays an array row. A general `row` function returning the whole row object would also carry a height, and later a hidden flag, at the price of an object per row for the common case; it can be added beside this one when a height per object is asked for.

## Errors and paths

A row that is neither form, a string, a number, `null`, an object without `cells`, is `INVALID_ROW` at the row; before, a string was written letter by letter and the others failed without a code. A height out of range is `INVALID_ROW_HEIGHT` at `rows[3].height`. A cell of a row object is named where it is, `rows[3].cells[2]`; a cell of an array row keeps `rows[3][2]`.

## What it costs

`Row` became a union, so code that reads `row.length` from a value typed `Row` narrows first; that is a change to the types and ships as a minor version. The three benchmark suites give their rows as arrays and measure within the noise band of `main`. The bundle grows by about 400 bytes gzipped; ADR 0021 raises the size budget for it.

## Considered and rejected

- `rowStyles` by row number: no use to a streamed sheet, and apart from the row it styles.
- The whole-row format in the file: reaches past the table, and leaves absent cells to each application.
- Extending a row's style to the sheet's column count: the reach would depend on something outside the row.
- A helper, `row(cells, { style })`: a second way to say what a plain object says.
- Sharing one loop between the two forms: the array path is the byte path, and stays as it was measured.
