# Padding around the picture of an image cell

An image cell takes `padding`: the space between the picture and each edge of its cell, in pixels. The row grows to the picture and its padding, and so does a column without a numeric width. It is optional, and a cell without it is written as before.

```ts
{ image: png, height: 40, padding: 10 }
```

The row is 60 pixels tall with the picture in its middle, and one picture is 20 pixels from the next down a column.

## Why

A picture sat two pixels inside the top-left corner of its cell and its row was the picture's height and four pixels, so a column of pictures was a column of pictures almost touching. That is right for a thumbnail beside a name and wrong where each picture has to be told from its neighbours: a code read by a scanner, a signature, a swatch. The caller could set a taller row, but the picture stayed in the corner.

The fix of ADR 0030 made this visible. Until then Excel for Mac drew every such row a third taller than written, and sheets looked spaced that were not.

## One number, on the cell

- **On the cell, beside `width` and `height`.** The space belongs to the picture, as its size does, so a column of pictures and a single logo say it the same way, and a `sheetFromObjects` value function returns it with the image. Nothing is added to the sheet or to the options.
- **Pixels**, the unit of `width` and `height`. A number from 0; fractions are taken as they are.
- **One number for all four sides.** The row needs the space above and below, the column left and right, and nobody has asked for them to differ. An object `{ x, y }` can be accepted later beside the number.
- **`padding: 0`** puts the picture in the corner and makes the row exactly its height.
- **Left out is not a padding of 2.** A cell without it keeps what ADR 0025 gave it: two pixels in from the corner, a row four pixels taller than the picture, and a column twelve pixels wider, the spare room a fitted column gives a text. Making that the same as `padding: 2` would narrow every picture column already written by eight pixels. So the default is "as before", and the padding is exact when asked for.

## What the writer does with it

The picture's anchor starts `padding` pixels in from the cell's top-left corner, where it started two pixels in. The row holding it takes the tallest of its pictures with the padding of each, in points at three quarters of a point per pixel; the column takes the widest, at seven pixels per character of width, which is the rule the file format gives for the default font. A picture in a hidden row or column keeps its two-cell anchor, moved in by the padding (ADR 0028).

The caller's own sizes win as they did: a height on the row, in `rowHeights` or on a row object, and a numeric column width stay as set, and the padding then only moves the picture in from the corner. A picture that needs more than a row can be with its padding, 545 pixels, is refused with `INVALID_IMAGE` at its cell unless the caller sizes the row, and a padding that is not a number from 0 is refused at `.padding`.

It is the same on both paths. The padding is on the cell, so a streamed row knows it when the row is written; the column follows the rule for picture columns when streaming (ADR 0025).

## How exact it is

The offsets and the row are exact wherever a point is a point. The column is exact where a character of the default font is seven pixels wide, as the format's rule has it; an application that draws a character wider shows more room right of the picture than left of it. Measured for `image-padding.xlsx`, in the compatibility log:

- Excel for Mac and Numbers draw a column of 15.71 characters 94 points wide, where the rule gives 82.5, so a picture with 7.5 points on its left has 19 on its right. Both show the row heights as written, Excel for Mac in whole points: 37 for 37.5.
- LibreOffice draws the same column 87 points wide.

No application measured draws a column narrower than the rule, so a picture and its padding do not reach into the next cell.

## Alignment through the style was considered

`align` and `verticalAlign` on the cell's style could have centred the picture. The writer would have to compute the offset from the width the application will draw the column at, which the measurements above show it cannot know, and from a row height that a streamed sheet may not have yet. A style also aligns the cell's content, and an image cell has none: the picture floats over the cell. The padding is a distance the writer writes down and every application reads the same way.

## Considered and rejected

- A sheet-level setting for all pictures: a second place to say what a cell holds, and a column of pictures gets it from one value function already.
- A default padding larger than the present two pixels: it would change every sheet already written.
- Rounding a picture row up to a whole point for Excel for Mac: half a point at most, and other applications would then show the row taller than asked.
- `margin`, `inset`, `spacing` as the name: `padding` is the word for space inside a box, and the cell is the box.
