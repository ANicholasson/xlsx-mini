# Sheet view: gridlines, zoom and right-to-left

A sheet can say how it is shown when it opens: `view: { gridlines?, zoom?, rightToLeft? }`. A report reads better without the lines between cells, a dense sheet at a smaller zoom, and a sheet in Arabic or Hebrew has column A at the right edge.

## One object, beside `freeze`

The three settings are attributes of the one `<sheetView>` element a worksheet part carries, the element the frozen pane already lives in. They are grouped as `view` on the sheet, so the sheet's own list of settings grows by one name and the next view setting has a home. `freeze` stays where it is: it has been on the sheet since 0.1.0, and moving it into `view` would break every caller for tidiness. The two are written into the same element and combine.

- `gridlines: false` hides the lines; `true`, the default of every application, writes nothing.
- `zoom` is a percentage, a whole number from 10 to 400, the range Excel accepts and the unit its status bar shows; 100 writes nothing. Only `zoomScale` is written. `zoomScaleNormal`, which Excel writes beside it, is the zoom to return to from the page-break preview; Excel and LibreOffice take the zoom without it.
- `rightToLeft: true` lays the sheet out from the right. It is about the sheet, not about a cell: the reading direction of a text in a cell is the application's own business and follows the text.

A sheet without a view, or with the defaults spelled out, writes the bytes it wrote before.

Update: every sheet now carries the bare sheet view, also one without a view or a frozen pane (ADR 0030). The settings and the pane are written into it as described here.

## A view is where the sheet starts

These are the state a sheet opens in, and the reader changes them freely. Each application keeps the settings it has a place for: the compatibility log for `sheet-view.xlsx` records which. Numbers has one zoom per window and ignores the file's, and Google Sheets opens at 100 % as well; whether `zoomScaleNormal` would change that in Sheets was not tried.

## Errors

The view is validated with the sheet's other settings, before any row: `INVALID_VIEW` at `sheets[0].view` for a view that is not an object, at `.zoom` for a zoom that is not a whole number from 10 to 400, at `.gridlines` and `.rightToLeft` for a value that is not a boolean. The flags the library had before, `autoFilter` and `autoWidth`, compare with `true` and are not checked; here a wrong type would silently show the sheet another way than asked, and the ticket asked for the check.

## Considered and rejected

- Three settings directly on the sheet, `gridlines`, `zoom`, `rightToLeft`: three more names in a list that is already long, and no home for the next one.
- Moving `freeze` into `view`: a breaking change for a grouping.
- A zoom as a factor, `1.5`: the applications show percent.
- Rounding a fractional zoom: 150.5 is a mistake to report, as a fractional frozen row count is.
- Writing `zoomScaleNormal` as Excel does: bytes for a preview the library does not write.
- The other view settings, the selected cell, row and column headers, zeros shown, the view kind: none asked for, each an attribute of the same element when one is.
