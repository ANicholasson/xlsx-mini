# Hidden rows, columns and sheets

A column, a row and a sheet can be hidden: `hidden: true` on a `Column`, on a row object and on a sheet. The uses are the ordinary ones: a sheet that only holds the source of a dropdown, an identifier column kept for a later import, a helper row. Hiding is for tidiness. The reader shows a hidden thing again from the application's menus, and the README says so where it describes the option.

## One word in three places

`hidden` sits where the thing it hides is described. A column and a sheet are known before any row, so theirs are settings. A row is hidden by the row itself, as it carries its style and its height (ADR 0027): a row from a generator or a cursor does not know its number, so a list of hidden row numbers on the sheet would be out of its reach. An array row cannot be hidden; it becomes `{ cells, hidden: true }`. `sheetFromObjects` takes `hidden` on a column and passes it through, the header cell going with its column.

Only `true` hides, as only `true` turns on `autoFilter` and `autoWidth`. A workbook that hides nothing writes the bytes it wrote before.

## Hiding changes what is seen and nothing else

A hidden row keeps its cells, its style and its height, and is written also when it has no cells, since the row is what is hidden. A hidden column keeps its width: a numeric one as given, a fitted one as measured, none when it has none, and Excel then shows it again at the default width. Formulas and dropdowns read hidden cells as they read any others.

The ticket for this change supposed that a hidden column needs no auto width. It is fitted all the same. Shown again it has the width its cells need; the rule is one sentence without an exception; and a fitted column in Excel counts the cells of hidden rows too, which was checked: a long text in a hidden row widened its column to the same 45 characters hidden and shown. A caller who minds the time gives the hidden column a number.

## The sheet a workbook opens on

Excel opens a workbook on its first sheet unless the workbook part names another, and a hidden sheet cannot be the one shown. When a hidden sheet comes first, the workbook part carries `<bookViews><workbookView activeTab="…"/></bookViews>` naming the first sheet that is not hidden; when the first sheet is shown, nothing is written. A workbook whose sheets are all hidden is one Excel refuses, so it fails with `NO_VISIBLE_SHEET` at `sheets` before anything is written, on both paths.

## A picture in a hidden row or column

A picture is anchored to its cell at its own size and never stretches with it (ADR 0025). That anchor, `oneCellAnchor`, has a cost here: with its row or column hidden the picture stays, lying across the cells that follow. Excel showed exactly that when it was tried, and a row hidden with its picture still in view is not hidden.

A picture in a hidden row or a hidden column therefore takes a `twoCellAnchor` with `editAs="twoCell"`, from the inset to the inset plus the picture, both ends inside its own cell. It is sized with the cell: nothing while the cell is hidden, its own size when the cell is shown again, which Excel confirmed for a row and for a column, hidden, shown and hidden again. The price is that such a picture, once shown, stretches when its row or column is dragged, which a picture elsewhere does not; and where the caller gave the column or the row a size smaller than the picture, it is squeezed into that size when shown. Every other picture keeps the anchor it had and its bytes.

The row and the column still make room for the picture, so that showing them gives the cell the picture needs.

## What the four applications do

Recorded in the compatibility log for `hidden.xlsx`. Excel and LibreOffice hide all three and open on the first sheet that is shown. Numbers has no hidden sheets: the hidden sheet is an ordinary one there, and the first, so the file opens on it; and its pictures float over the table, so one in a hidden row or column stays in view. Neither can be changed from the file, and both are in the README beside the table.

## Room for what is not here

- **Very hidden**, the sheet state only a macro or the file itself can undo: `hidden` on a sheet can later take `"very"` beside `true`.
- **Outline groups**, with their levels and their collapsed state: an `outlineLevel` and a `collapsed` flag on a column and on a row object, beside `hidden`, which a collapsed group sets on its members.

Neither is asked for yet, and neither needs a field renamed.

## Considered and rejected

- `hiddenRows` by row number on the sheet: no use to a streamed sheet, and apart from the row it hides.
- `visibility: "visible" | "hidden"` on the sheet: a second word for what the column and the row call `hidden`.
- A width of 0 for a hidden column without one, as some writers store it: the column then has no width to come back to in every application.
- Leaving a hidden column unmeasured: an exception to the rule above for a saving of one column's measuring.
- Leaving pictures as they were and writing down that they stay in view: the option would not do what its name says.
- Refusing an image cell in a hidden row or column: a thumbnail column a caller hides is a reasonable workbook.
- Naming the active sheet in every workbook: bytes in every file for a default every application already has.
