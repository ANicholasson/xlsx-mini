# Images as a cell type, anchored to their cell

A cell can hold a PNG: `{ image: bytes }`, with an optional `width` and `height` in pixels and a `style` for the cell beneath. The picture is written once per distinct image as a media part and anchored to its cell in the sheet's drawing part, at its pixel size, moving with the cell and never stretching with it. The row takes the picture's height and, where the column has no numeric width, the column takes its width. Producing the image stays with the caller; the writer reads the PNG header for the size and nothing else. A workbook without images writes the same bytes as before. This closes the last item ADR 0019 deferred.

## One more kind of cell, not a new surface

The library has one way to say what a cell holds: a value, or a plain object that names its kind, `value`, `formula`, `href`, `cell`. An image is `{ image }` in the same place, so a column of pictures comes out of the rows API that already exists, and a `sheetFromObjects` value function returns it like any other cell. That function's return type widens from a bare value to a `Cell`, which also lets it return a styled or linked cell, a gap in the 0.1.0 shape. No sheet-level list of images, no option on `writeXlsx`: a picture belongs to a cell the way a formula does, and the renderer records it while the row is being written, as it records links.

An image cell has no `value`; the cell beneath the picture stays empty, and `style` reaches it for a fill or a border. An image cannot also be a formula or a link, and the error is `INVALID_IMAGE` at the cell. Bare bytes in a row are not an image: the kind is named, so a later format needs no guessing, and a `Uint8Array` that reached a cell by mistake stays an `UNSUPPORTED_VALUE`.

## What is read from the image, and what is written

- **PNG only, by its header.** The eight-byte signature and the `IHDR` chunk give the pixel size; anything else is `INVALID_IMAGE` with the reason. JPEG is a later addition with the same shape, since its header needs a scan for the frame marker; neither needs a decoder, and the library stays without a dependency.
- **Size in pixels.** `width` and `height` default to the header; one given scales the other by the aspect ratio, both given are taken as they are. Pixels because that is the unit the caller sized the image in; the file stores EMUs at 9 525 per pixel, Excel's 96 dpi.
- **A one-cell anchor.** `<xdr:oneCellAnchor>` from the cell's top-left, inset two pixels, with the picture's extent. A two-cell anchor stretches the picture with its cell, which distorts a photo and can stop a barcode scanning; a one-cell anchor moves with the cell and keeps its size, which is what a picture in a table should do.
- **One media part per distinct image.** Images are keyed by length and CRC-32 and compared byte for byte on a match, so a logo repeated down a column is one `xl/media/imageN.png` and many anchors. Media parts are numbered across the workbook, since the same image may appear on several sheets.
- **One drawing part per sheet that has images**, `xl/drawings/drawingN.xml` with its own relationships to the media, a relationship from the sheet to the drawing, and `<drawing r:id>` in the sheet after the print setup, where the schema puts it. The content types gain a default for `png` and an override per drawing part.

## Sizing: the row always, the column when it may

A picture lying across neighbouring cells is a broken table, so the writer sizes what it can:

- **The row** holding the image gets the picture's height plus the inset, in points at three quarters of a point per pixel and never less than the default row, unless `rowHeights` names that row, in which case the caller's height wins. A row with several images takes the tallest. Excel caps a row at 409 points, 541 pixels of picture; a taller picture whose row the caller has not sized is refused rather than left to run into the next row. This happens on both paths, since a row's height is written with the row.
- **The column** holding the image is measured like text under auto width: an image column with `width: "auto"`, or under `autoWidth`, counts the picture's width among its cells; an image column with no width at all is widened to its widest picture, since the default width fits no picture worth embedding. A numeric width is the caller's and stays. Column widths precede the rows in the file, so when streaming this follows the auto-width rule of ADR 0024: an image column is sized from the first 1 000 rows when the sheet holds the head for auto width, and otherwise the caller sets the width, which the README says next to the option.

Nothing here is particular to any kind of picture; it is what a cell that holds one needs.

## Where the parts go, on both paths

Anchors are recorded as rows are rendered, so the sheet is read once on both paths. The drawing part and the sheet's new media parts are written right after the sheet part, as the relationships part already is. On the byte path everything is known before packaging. When streaming, the images of one sheet are held until that sheet ends, deduplicated, because the parts that carry them follow the sheet; memory grows with the distinct images of the sheet in hand, not with its rows, and the README says so where it describes streaming.

The content types part names every drawing part and the `png` default, and a streamed package cannot know before the first sheet whether any sheet will have images. The streamed package therefore writes `[Content_Types].xml` last. ADR 0022 already moved the workbook, styles and shared-strings parts after the sheets on the ground that readers locate parts through the central directory, and this is the same move; the streamed fixtures are opened again in the four applications to show it holds for this part too. The byte path keeps its order and its bytes.

## What the change must also carry

A `drawing` feature slice, which the dependency rules have reserved since ADR 0013, holding the media table, the per-sheet drawing and the PNG header reader, behind one entry. A golden fixture with several images of two sizes, one of them repeated, a header row and text columns, opened in the four applications and recorded in the compatibility log, with the check that no picture overlaps another cell. The streamed fixture identity for it. A README section with a snippet run by the tests, the API and error tables, a glossary entry, an update note in ADR 0019, and a minor changeset. The benchmark suites hold no images and must stay inside the noise band.

## Considered and rejected

- A sheet-level `images: [{ cell, bytes }]` list: a second way to say what a cell holds, out of reach of `sheetFromObjects` and of a row generator, and anchors that cannot be recorded while rows stream.
- Bytes as a cell value without a wrapper: ambiguous once a second format exists, and a stray byte array would silently become a picture.
- Two-cell anchors stretched to the cell, as some writers do: pictures that distort when a column is dragged.
- Excel's in-cell pictures (rich values): Excel-only, a second metadata part and a rich-value store, invisible to the other three applications.
- Decoding image data for a size the header already states.
- Declaring the `png` default and drawing overrides in every streamed package up front: an override for a part that does not exist is a malformed package, and a default nobody uses is a lie about every file.
