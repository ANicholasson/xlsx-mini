# Every sheet carries a sheet view

Every worksheet part is written with `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`, also when the sheet asks for no view and freezes nothing. Without the element, Excel for Mac shows every row that has a height a third taller than the height written.

## What was found

A workbook written by 0.8.0 with row heights of 30 and 48 points read 40 and 64 in Excel for Mac 16.115, on screen and through AppleScript. Rows without a height, column widths and the size of a picture were as written, so a picture sized to fill its row stopped three quarters of the way down. Numbers and LibreOffice read the same file at the heights written.

The file was patched one element at a time and read back after each change. The sheet view decided it, and nothing else did:

- With `<sheetViews><sheetView workbookViewId="0"/></sheetViews>` the rows read 30 and 48. The same holds for any view the writer already wrote: a frozen pane, hidden gridlines, a zoom.
- No effect: the workbook's file version, properties and book views; the sheet's default row height and base column width; `x14ac:dyDescent` on the rows; the default font and its size; the font scheme.
- It is per sheet. In a workbook of three sheets, only those with a sheet view read right.

Why Excel does it is not known. Every file Excel writes itself carries a sheet view, so a part without one is a file it never sees from its own hand. The compatibility log had recorded the effect twice, for `appearance.xlsx` and `row-styles.xlsx`, as a quirk of one machine; those rows are corrected.

## The decision

The sheet view is written always. It is one bare element, the form the schema allows and Excel writes for a sheet nobody has touched, and the settings of `view` and the pane of `freeze` go into it as before (ADR 0029).

It is not written only for sheets that need it, because the writer cannot know which do. A streamed sheet writes its head before its first row, and a row with a height or a picture may come at any point after. A sheet whose rows all keep the default height would read the same without the element, but two forms of the same part for no visible gain is one more thing to test.

The cost is 55 bytes per sheet before compression, and no code: the branch that left the element out is gone.

## What changes for a caller

Nothing in the API. Every workbook's bytes change by the one element per sheet. In Excel for Mac, a row with a height now shows at that height, so a sheet that looked right there only because its rows came out taller than written looks tighter; the remedy is to write the height that is wanted.

## Not checked

Excel on Windows. Whether it shares the behaviour is unknown; the fix does not depend on the answer, since the element is one Excel writes itself.

## Considered and rejected

- Writing the element only when a sheet sets a row height or holds a picture: not knowable when streaming, and two forms of one part.
- Writing three quarters of each height: it would correct Excel for Mac and break every other application.
- An option to turn the element off: nobody is helped by a file that reads wrong in Excel.
