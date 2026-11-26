# xlsx-mini

[![npm](https://img.shields.io/npm/v/xlsx-mini)](https://www.npmjs.com/package/xlsx-mini)
[![ci](https://github.com/ANicholasson/xlsx-mini/actions/workflows/ci.yml/badge.svg)](https://github.com/ANicholasson/xlsx-mini/actions/workflows/ci.yml)
[![size](https://img.shields.io/bundlejs/size/xlsx-mini?exports=writeXlsx)](https://bundlejs.com/?q=xlsx-mini&treeshake=[{writeXlsx}])
[![license](https://img.shields.io/npm/l/xlsx-mini)](./LICENSE)

Writes `.xlsx` files from TypeScript with no dependencies, on every JavaScript runtime. You describe sheets, rows and cells; it returns the bytes. An experimental entry, `xlsx-mini/read`, reads the values of an `.xlsx` file and refuses one that is damaged or hostile. The files open without a repair dialog in Excel, Numbers, Google Sheets and LibreOffice, and the writer is measured on every change. Every TypeScript snippet in this README is run by the test suite.

## Install

```sh
pnpm add xlsx-mini
```

ESM only. Node 20.12 or later (22 or later for `require`), Bun, Deno, Cloudflare Workers, Vercel's edge runtime and evergreen browsers: the core uses only web-standard APIs, and the runtime matrix in CI proves each host on every change.

In a bundle, minified and gzipped, an application that imports only `writeXlsx` takes on under 19 kB, and every export together under 22 kB. A test holds both on every change, and the badge above is the same measurement of the version on npm by a service outside this repository.

## Quick start

<!-- snippet: quick-start -->

```ts
import { writeFile } from "node:fs/promises";
import { writeXlsx } from "xlsx-mini";

export default async function quickStart(path = "orders.xlsx"): Promise<void> {
  const bytes = await writeXlsx({
    sheets: [
      {
        name: "Orders",
        rows: [
          ["Order", "Customer", "Total", "Paid"],
          ["A-1001", "Kari Nordmann", 1234.5, true],
          ["A-1002", "Ola Nordmann", 99, false],
        ],
      },
    ],
  });
  await writeFile(path, bytes);
}
```

A row is an array of cells. A cell is a string, a number, a boolean, a `Date`, a Temporal value, or `null` for an empty cell; wrap a value in `{ value, style }` to style it. The result is an ArrayBuffer-backed `Uint8Array`, ready for a file, a `Response` or a `Blob`.

## Where it runs

### Any web framework, through a Response

<!-- snippet: response -->

```ts
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

export default async function report(): Promise<Response> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  return new Response(bytes, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="report.xlsx"',
    },
  });
}
```

### Express and friends

<!-- snippet: express -->

```ts
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

/** Any response with setHeader and end: Express, Node's http, Fastify's raw reply. */
interface Reply {
  setHeader(name: string, value: string): void;
  end(body: Uint8Array): void;
}

export default async function handler(reply: Reply): Promise<void> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  reply.setHeader("Content-Type", XLSX_MIME_TYPE);
  reply.setHeader("Content-Disposition", 'attachment; filename="report.xlsx"');
  reply.end(bytes);
}
```

### Browser

The `xlsx-mini/browser` entry adds `downloadXlsx`, which saves the bytes through a download. It throws `NO_DOM` anywhere without a document, so a server never downloads by accident.

<!-- snippet: browser -->

```ts
import { writeXlsx } from "xlsx-mini";
import { downloadXlsx } from "xlsx-mini/browser";

export default async function exportInBrowser(): Promise<void> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  downloadXlsx(bytes, "report.xlsx");
}
```

### Next.js route handler

<!-- snippet: next-route -->

```ts
// app/api/report/route.ts
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

export async function GET(): Promise<Response> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  return new Response(bytes, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="report.xlsx"',
    },
  });
}
```

The same handler on the edge runtime:

<!-- snippet: next-edge -->

```ts
// app/api/report/route.ts on the edge runtime: the same handler, one line more
import { XLSX_MIME_TYPE, writeXlsx } from "xlsx-mini";

export const runtime = "edge";

export async function GET(): Promise<Response> {
  const bytes = await writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
  return new Response(bytes, { headers: { "Content-Type": XLSX_MIME_TYPE } });
}
```

### Next.js server action and client component

The action builds the workbook on the server; the client receives the bytes and downloads them.

<!-- snippet: actions -->

```ts
// app/actions.ts
"use server";

import { writeXlsx } from "xlsx-mini";

export async function exportReport(): Promise<Uint8Array> {
  return writeXlsx({ sheets: [{ name: "Report", rows: [["Total", 42]] }] });
}
```

<!-- snippet: next-client -->

```ts
// app/export-button.tsx: wire this to a button's onClick
"use client";

import { downloadXlsx } from "xlsx-mini/browser";
import { exportReport } from "./actions";

export async function onExportClick(): Promise<void> {
  downloadXlsx(await exportReport(), "report.xlsx");
}
```

## Streaming large workbooks

`streamXlsx` returns a web `ReadableStream` of the file's bytes, produced as the consumer reads them, and takes rows from an async iterable, so a database cursor can feed a sheet page by page. Peak memory stays flat however many rows go out. Return the stream in a `Response` from any route handler, pipe it to a file, or `for await` over its chunks.

<!-- snippet: streaming -->

```ts
import { type Row, XLSX_MIME_TYPE, streamXlsx } from "xlsx-mini";

/** Rows arrive page by page from wherever they live; nothing is held beyond the page in hand. */
async function* orders(fetchPage: (page: number) => Promise<Row[]>): AsyncGenerator<Row> {
  yield ["Order", "Customer", "Total"];
  for (let page = 0; ; page++) {
    const rows = await fetchPage(page);
    if (rows.length === 0) return;
    yield* rows;
  }
}

/** The response a route handler returns; `fetchPage` is however the application reads a page. */
export default function ordersResponse(fetchPage: (page: number) => Promise<Row[]>): Response {
  const stream = streamXlsx({
    sheets: [{ name: "Orders", rows: orders(fetchPage), autoFilter: true }],
  });
  return new Response(stream, {
    headers: {
      "Content-Type": XLSX_MIME_TYPE,
      "Content-Disposition": 'attachment; filename="orders.xlsx"',
    },
  });
}
```

Text is written into each cell rather than through a shared-strings table, which is what keeps memory flat; `strings: "shared"` keeps the table instead, for a smaller file at the cost of holding every distinct string until the end. The stream throws before it exists when the workbook itself is invalid, and errors mid-way when a cell cannot be written. Cancelling it, as a client going away does, closes the row source. A column with `width: "auto"` is sized from the first 1 000 rows, since the widths go out before the rows, and a sheet's images are held until the sheet ends.

## From an array of objects

`sheetFromObjects` turns objects into a sheet: a bold header row, a frozen header and an autofilter by default, each switchable. A column names a key, and the types reject a key whose values a cell cannot hold; a `value` function computes a cell from the whole object.

<!-- snippet: objects -->

```ts
import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Order {
  id: string;
  customer: string;
  total: number;
  placed: Date;
}

export default function fromObjects(orders: Order[]): Promise<Uint8Array> {
  const sheet = sheetFromObjects(orders, {
    name: "Orders",
    columns: [
      { key: "id", header: "Order", width: 12 },
      { key: "customer", header: "Customer", width: 24 },
      { key: "total", header: "Total", width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
      { key: "placed", header: "Placed", width: 18, style: { numberFormat: "yyyy-mm-dd hh:mm" } },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
```

## Auto column width

`width: "auto"` on a column fits it to its widest cell; `autoWidth: true` on a sheet does that for every column without a numeric width. The width is estimated from what each cell shows: text character by character against Calibri 11, numbers and dates by the length their format produces, bold a little wider, from 3 to 60 characters. No font is measured, so the fit is close rather than exact and errs wide; a column whose cells are all empty keeps the application's default. When streaming, the first 1 000 rows of a sheet set its widths and the rows after them take those widths, because the widths go out before the rows (ADR 0024).

<!-- snippet: auto-width -->

```ts
import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Product {
  sku: string;
  name: string;
  price: number;
  restocked: Date;
  note: string;
}

export default function inventory(products: Product[]): Promise<Uint8Array> {
  // Every column without a numeric width fits its widest cell, bold header included; the note
  // column keeps the width it is given. On a single column, width: "auto" does the same.
  const sheet = sheetFromObjects(products, {
    name: "Inventory",
    autoWidth: true,
    columns: [
      { key: "sku", header: "SKU" },
      { key: "name", header: "Product" },
      { key: "price", header: "Unit price", style: { numberFormat: '#,##0.00 "kr"' } },
      { key: "restocked", header: "Restocked", style: { numberFormat: "yyyy-mm-dd" } },
      { key: "note", header: "Note", width: 40, style: { wrap: true } },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
```

## Number formats

A style sets `bold` and a `numberFormat`: an Excel format string, or a built-in format id from 0 to 49; the next section covers fonts, fills, borders and alignment. Column styles apply to every cell in the column; a cell's own style wins.

<!-- snippet: formats -->

```ts
import { type CellStyle, writeXlsx } from "xlsx-mini";

const header: CellStyle = { bold: true };
const money: CellStyle = { numberFormat: '#,##0.00 "kr"' };
const percent: CellStyle = { numberFormat: "0.0%" };
const isoDate: CellStyle = { numberFormat: "yyyy-mm-dd" };

export default function formats(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Invoice",
        // A column style applies to every cell in the column; a cell's own style wins.
        columns: [undefined, { style: money }, { style: percent }, { style: isoDate }],
        rows: [
          [
            { value: "Item", style: header },
            { value: "Price", style: header },
            { value: "VAT", style: header },
            { value: "Due", style: header },
          ],
          ["Desk", 4999, 0.25, new Date(Date.UTC(2026, 11, 1))],
          [
            "Chair",
            { value: 1299.5, style: { numberFormat: "0" } },
            0.25,
            new Date(Date.UTC(2026, 11, 15)),
          ],
        ],
      },
    ],
  });
}
```

| Want                  | Format string      |
| --------------------- | ------------------ |
| Thousands, 2 decimals | `#,##0.00`         |
| Currency after        | `#,##0.00 "kr"`    |
| Currency before       | `"$"#,##0.00`      |
| Percent               | `0.0%`             |
| Whole number          | `0`                |
| ISO date              | `yyyy-mm-dd`       |
| Date and time         | `yyyy-mm-dd hh:mm` |
| Text as typed         | `@`                |

## Fonts, fills, borders and alignment

A style also sets the font (`italic`, `fontSize` in points, `fontColor`), a solid `fill`, a `border` for all four sides or per side (`thin`, `medium`, `thick`), `align`, `verticalAlign`, `wrap` and `indent`. An indent aligns left unless `align` is `right`. Colours are `#rrggbb`. A sheet merges ranges with `merges` and sets `rowHeights` in points by row number; a row can also carry its own height and a style, as the next section shows. A cell object with a style and a `null` value writes an empty styled cell, which is how a border or a fill runs across a blank.

<!-- snippet: appearance -->

```ts
import { type CellStyle, writeXlsx } from "xlsx-mini";

const title: CellStyle = {
  bold: true,
  fontSize: 16,
  fontColor: "#ffffff",
  fill: "#1f4e79",
  align: "center",
  verticalAlign: "middle",
};
const heading: CellStyle = { bold: true, fill: "#d9e1f2", border: "thin" };
const total: CellStyle = { bold: true, border: { top: "medium", bottom: "thick" } };

export default function appearance(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Invoice",
        columns: [
          { width: 14 },
          { width: 36 },
          { width: 8, style: { align: "right" } },
          { width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
        ],
        // The title spans four columns; its row and the wrapped row get a height in points.
        merges: ["A1:D1"],
        rowHeights: { 1: 28, 3: 48 },
        rows: [
          [{ value: "Invoice 2026-1042", style: title }],
          [
            { value: "Item", style: heading },
            { value: "Description", style: heading },
            { value: "Qty", style: heading },
            { value: "Amount", style: heading },
          ],
          [
            "Desk",
            {
              value: "Height-adjustable desk, delivered and assembled on site",
              style: { wrap: true, verticalAlign: "top" },
            },
            2,
            9998,
          ],
          ["Chair", { value: "Mesh office chair", style: { indent: 1 } }, 4, 5198],
          [
            { value: "Total", style: total },
            // A styled cell without a value: the rule runs across the blank.
            { value: null, style: total },
            { value: 6, style: total },
            { value: 15196, style: total },
          ],
          [],
          [{ value: "Payment within 30 days.", style: { italic: true, fontColor: "#808080" } }],
        ],
      },
    ],
  });
}
```

Every distinct style is written once, however many cells share it, so styling costs nothing per cell beyond a lookup.

## Row styles

A row is an array of cells, or an object `{ cells, style, height }` when the row itself carries something. The style reaches every position in `cells`, the empty ones included, so a fill or a border runs across a blank, and it stops where the row stops. A cell's own style wins over its row's, and the row's over its column's, property by property: a bold total row keeps its column's number format. The height is in points and wins over `rowHeights`, which a row from a generator or a cursor cannot use, since it does not know its number. Give the rows that share a style the same style object.

<!-- snippet: row-styles -->

```ts
import { type CellStyle, type Row, writeXlsx } from "xlsx-mini";

const header: CellStyle = { bold: true, fill: "#1f2937", fontColor: "#ffffff" };
const band: CellStyle = { fill: "#f3f4f6" };
const total: CellStyle = { bold: true, border: { top: "thin" } };

type Order = [id: string, customer: string, amount: number];

export default function report(orders: Order[]): Promise<Uint8Array> {
  const sum = orders.reduce((soFar, [, , amount]) => soFar + amount, 0);
  const rows: Row[] = [
    { cells: ["Order", "Customer", "Amount"], style: header, height: 24 },
    // Every second order on a band; the others stay plain arrays.
    ...orders.map((order, index): Row => (index % 2 === 1 ? { cells: order, style: band } : order)),
    // The rule above the total runs across the empty cell.
    { cells: ["Total", null, sum], style: total },
  ];
  return writeXlsx({
    sheets: [
      {
        name: "Report",
        columns: [{ width: 12 }, { width: 24 }, { width: 14, style: { numberFormat: "#,##0.00" } }],
        rows,
      },
    ],
  });
}
```

`sheetFromObjects` takes `rowStyle`, a function from an object and its position to a style or `undefined`, for a band or for the rows that need attention.

## Hidden rows, columns and sheets

`hidden: true` on a column, on a row object or on a sheet takes it out of sight. The reader can show it again from the application's menus, so it tidies a workbook and keeps nothing secret. What is hidden stays in the file as it was: formulas and dropdowns read it, a hidden column keeps its width and a hidden row its height, and a picture there folds away with its row or column. A workbook needs one sheet that is not hidden, and opens on the first of them.

<!-- snippet: hidden -->

```ts
import { type Row, writeXlsx } from "xlsx-mini";

type Order = { id: number; number: string; status: string; amount: number; cancelled: boolean };

const statuses = ["Open", "Shipped", "Closed"];

export default function orderSheet(orders: Order[]): Promise<Uint8Array> {
  const rows: Row[] = [
    ["Id", "Order", "Status", "Amount"],
    ...orders.map((order): Row => {
      const cells = [order.id, order.number, order.status, order.amount];
      // A cancelled order stays in the file and in the sums, out of sight.
      return order.cancelled ? { cells, hidden: true } : cells;
    }),
  ];
  return writeXlsx({
    sheets: [
      {
        name: "Orders",
        // The id is kept for a later import, not for the reader.
        columns: [{ hidden: true }, { width: 12 }, { width: 12 }, { width: 14 }],
        rows,
        validations: [{ range: "C2:C1000", list: `Lists!A1:A${statuses.length}` }],
      },
      // The dropdown reads its values from a sheet nobody needs to see.
      { name: "Lists", hidden: true, rows: statuses.map((status) => [status]) },
    ],
  });
}
```

`sheetFromObjects` takes `hidden` on a column, which hides its header with it.

## Sheet view

`view` on a sheet says how it is shown when it opens. `gridlines: false` hides the lines between cells, `zoom` is a percentage from 10 to 400, and `rightToLeft: true` lays the sheet out from the right, column A at the right edge, as Arabic and Hebrew read. The frozen panes of `freeze` are part of the same view, so the two combine. A view is where the sheet starts: the reader changes it freely, and Numbers and Google Sheets open at their own zoom whatever the file says.

<!-- snippet: sheet-view -->

```ts
import { writeXlsx } from "xlsx-mini";

export default function report(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Report",
        // A report reads better without the grid, a little larger.
        view: { gridlines: false, zoom: 120 },
        rows: [
          ["Quarter", "Revenue"],
          ["Q1", 120000],
          ["Q2", 135500],
        ],
      },
      {
        name: "فواتير",
        // Column A sits at the right edge; the frozen header row is part of the same view.
        view: { rightToLeft: true },
        freeze: { rows: 1 },
        rows: [
          ["الطلب", "المبلغ"],
          ["A-1001", 1234.5],
        ],
      },
    ],
  });
}
```

## Formulas and hyperlinks

A cell `{ formula: "SUM(B2:B4)" }` holds a formula, written without the leading `=`. The applications calculate it when the file opens; a cached `value` shows until then and is optional. Formulas are escaped and never evaluated, and text that starts with `=` stays text, so data never becomes code. Functions newer than Excel 2007, such as `XLOOKUP`, `IFS` and `TEXTJOIN`, are stored under the prefix the file format wants, so a formula works as you type it; one that calls `LET` or `LAMBDA` is written exactly as given, and a function that returns an array shows its first value rather than spilling. Google Sheets does not know the name Excel stores `FILTER` and `SORT` under, so those two show `#NAME?` there. A cell `{ value, href }` links its text to an absolute URL, `{ value, cell }` jumps to a cell or range on the same sheet, and `{ value, sheet, cell }` to another sheet. A link takes no look of its own: `underline` and a `fontColor` make it look like one.

<!-- snippet: formulas -->

```ts
import { type CellStyle, writeXlsx } from "xlsx-mini";

const link: CellStyle = { fontColor: "#0563c1", underline: true };

export default function formulasAndLinks(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Budget",
        rows: [
          ["Item", "Qty", "Price", "Total"],
          ["Desk", 2, 4999, { formula: "B2*C2" }],
          ["Chair", 4, 1299.5, { formula: "B3*C3" }],
          // A cached value shows until the application recalculates; it is optional.
          ["Sum", { formula: "SUM(B2:B3)", value: 6 }, null, { formula: "SUM(D2:D3)" }],
          [],
          [{ value: "Price list", href: "https://example.com/prices", style: link }],
          [{ value: "Go to the summary", sheet: "Summary", cell: "A1", style: link }],
        ],
      },
      { name: "Summary", rows: [["Grand total", { formula: "Budget!D4" }]] },
    ],
  });
}
```

## Images

A cell can hold a PNG: `{ image: bytes }`, with an optional `width` and `height` in pixels, one scaling the other, and a `style` for the cell beneath. The picture is anchored to its cell at that size and moves with it without stretching; the row grows to the picture's height, and so does a column without a numeric width. Identical images are written once however many cells show them. The writer reads only the PNG header; producing the image is the application's job.

<!-- snippet: images -->

```ts
import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Product {
  name: string;
  code: string;
  /** The bytes of a PNG the application produced; the writer reads only its size. */
  picture: Uint8Array;
}

export default function catalogue(products: Product[]): Promise<Uint8Array> {
  // A picture is a cell like any other: the row grows to its height and the column to its width.
  const sheet = sheetFromObjects(products, {
    name: "Catalogue",
    autoWidth: true,
    columns: [
      { key: "name", header: "Product" },
      { key: "code", header: "Code" },
      {
        key: "picture",
        header: "Picture",
        value: (product) => ({ image: product.picture, height: 40 }),
      },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
```

`padding` is the space between the picture and each edge of its cell, in pixels. The row grows to the picture and its padding, so the picture sits in the middle of its row, and with `padding: 10` one picture is 20 pixels from the next down a column. A column without a numeric width is at least as wide as the picture and its padding, and the picture starts its padding in from the left. Without `padding` the picture sits two pixels inside the top-left corner of its cell. A row height or a column width the caller sets stays as set; the padding then only moves the picture in from the corner.

<!-- snippet: image-padding -->

```ts
import { sheetFromObjects, writeXlsx } from "xlsx-mini";

interface Item {
  name: string;
  /** The bytes of a PNG the application produced, such as a code to scan. */
  label: Uint8Array;
}

export default function labels(items: Item[]): Promise<Uint8Array> {
  const sheet = sheetFromObjects(items, {
    name: "Labels",
    autoWidth: true,
    columns: [
      { key: "name", header: "Item" },
      {
        key: "label",
        header: "Label",
        // 10 pixels on every side: each row is 60 pixels tall, and one label is 20 from the next.
        value: (item) => ({ image: item.label, height: 40, padding: 10 }),
      },
    ],
  });
  return writeXlsx({ sheets: [sheet] });
}
```

When streaming, the images of a sheet are held until the sheet ends, since the parts that carry them follow it; set the image column's width or use `autoWidth`, because column widths go out before the rows.

## Data validation

A sheet's `validations` say what the cells of a range accept: a dropdown from a list of values or from a range of cells, a number or a date within bounds, each with an optional input message and error message. A range is a cell, a range in A1 form, or whole columns such as `E:E`. The rules are declared on the sheet, so a sheet whose rows come from a generator needs no second pass.

<!-- snippet: validation -->

```ts
import { type DataValidation, writeXlsx } from "xlsx-mini";

const status: DataValidation = {
  range: "B2:B1000",
  list: ["Open", "Closed", "On hold"],
  input: { title: "Status", message: "Pick one from the list" },
};
const quantity: DataValidation = {
  range: "C2:C1000",
  number: { min: 1, max: 1000, integer: true },
  error: { title: "Quantity", message: "A whole number from 1 to 1000" },
};
// A dropdown read from cells, here a column on another sheet.
const owner: DataValidation = { range: "D2:D1000", list: "People!A2:A50" };
// Whole columns are written as E1:E1048576; a warning lets the entry through after a prompt.
const due: DataValidation = {
  range: "E:E",
  date: { min: new Date(Date.UTC(2026, 0, 1)) },
  error: { message: "Due dates fall in 2026 or later", style: "warning" },
};

export default function validated(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Orders",
        rows: [
          ["Order", "Status", "Qty", "Owner", "Due"],
          ["A-1001", "Open", 2, "Kari", new Date(Date.UTC(2026, 11, 1))],
        ],
        validations: [status, quantity, owner, due],
      },
      { name: "People", rows: [["Name"], ["Kari"], ["Ola"]] },
    ],
  });
}
```

## Print setup

A sheet's `print` sets the orientation, the paper size by name (`letter`, `legal`, `tabloid`, `A3`, `A4`, `A5`), fit to a number of pages wide or tall, margins in inches as Excel stores them, header and footer text with Excel's codes, and `repeatRows`, the header rows printed at the top of every page.

<!-- snippet: print -->

```ts
import { type PrintSetup, writeXlsx } from "xlsx-mini";

// &L, &C and &R start the left, centre and right sections; &P, &N, &D and &A are the page
// number, page count, date and sheet name; && is an ampersand.
const print: PrintSetup = {
  orientation: "landscape",
  paperSize: "A4",
  fitTo: { width: 1 },
  margins: { left: 0.5, right: 0.5 },
  header: "&LQuarterly report&R&D",
  footer: "&L&A&RPage &P of &N",
  repeatRows: 1,
};

export default function printable(): Promise<Uint8Array> {
  const rows = Array.from({ length: 200 }, (_, i) => [
    i + 1,
    `Region ${(i % 4) + 1}`,
    (i + 1) * 125,
  ]);
  return writeXlsx({
    sheets: [{ name: "Report", rows: [["Week", "Region", "Revenue"], ...rows], print }],
  });
}
```

## Dates and Temporal

A `Date` is written as its UTC instant; what the reader shows follows the cell's format, or the built-in date-time format when the cell has none. Temporal values need no polyfill: `Temporal.PlainDate`, `PlainDateTime`, `ZonedDateTime` and `Instant` are recognised by their `Symbol.toStringTag`, and a `PlainDate` gets the built-in date format. Dates run from 1900-01-01 to 9999-12-31, in Excel's 1900 date system.

<!-- snippet: dates -->

```ts
import { writeXlsx } from "xlsx-mini";

export default function dates(): Promise<Uint8Array> {
  return writeXlsx({
    sheets: [
      {
        name: "Dates",
        rows: [
          // A Date is written as its UTC instant and shown with the built-in date-time format.
          ["Shipped", new Date("2026-11-05T14:30:00Z")],
          // A Temporal.PlainDate shows as a date; PlainDateTime, ZonedDateTime and Instant as
          // date-times. They are recognised by their Symbol.toStringTag, so no polyfill is needed.
          ["Due", { [Symbol.toStringTag]: "Temporal.PlainDate", year: 2026, month: 12, day: 1 }],
        ],
      },
    ],
  });
}
```

## Extending accepted value types

Augment the `Register` interface with your own types, and the `coerce` option converts each such value once, with its path, into something a cell can hold. `coerce` also sees every `Date` and Temporal value, so it can override how those are written.

<!-- snippet: register -->

```ts
import { type Coerce, writeXlsx } from "xlsx-mini";

export class Money {
  constructor(readonly cents: number) {}
}

// Let your own types into cells; the types then accept them everywhere a cell value goes.
declare module "xlsx-mini" {
  interface Register {
    cellValue: Money | bigint;
  }
}

// Then say how they become something a cell can hold.
const coerce: Coerce = (value) => {
  if (value instanceof Money) return value.cents / 100;
  if (typeof value === "bigint") return Number(value);
  return value;
};

export default function extended(): Promise<Uint8Array> {
  return writeXlsx(
    {
      sheets: [
        {
          name: "Ledger",
          rows: [
            ["Rent", new Money(1_250_000)],
            ["Views", 1234567890123n],
          ],
        },
      ],
    },
    { coerce },
  );
}
```

## Errors and limits

The writer throws one error class, `XlsxMiniError`, with a `code` to switch on and a `path` into your input such as `sheets[1].rows[40][2]`. It throws before writing anything a reader would refuse.

<!-- snippet: errors -->

```ts
import { XlsxMiniError, writeXlsx } from "xlsx-mini";

export default async function safely(): Promise<string> {
  try {
    await writeXlsx({ sheets: [{ name: "Bad/Name", rows: [] }] });
    return "written";
  } catch (error) {
    // code says what, path says where: "INVALID_SHEET_NAME at sheets[0].name"
    if (error instanceof XlsxMiniError) return `${error.code} at ${error.path}`;
    throw error;
  }
}
```

| Code                    | When                                                                                                                                                                     |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `NO_SHEETS`             | The workbook has no sheet.                                                                                                                                               |
| `NO_VISIBLE_SHEET`      | Every sheet of the workbook is hidden.                                                                                                                                   |
| `INVALID_SHEET_NAME`    | Empty, over 31 characters, contains `[ ] : * ? / \`, starts or ends with `'`, or is `History`.                                                                           |
| `DUPLICATE_SHEET_NAME`  | Two sheets share a name, ignoring case.                                                                                                                                  |
| `INVALID_NUMBER`        | `NaN` or an infinity.                                                                                                                                                    |
| `INVALID_DATE`          | An invalid `Date`.                                                                                                                                                       |
| `DATE_OUT_OF_RANGE`     | Before 1900-01-01 or after 9999-12-31.                                                                                                                                   |
| `UNSUPPORTED_VALUE`     | A value no cell can hold and `coerce` did not convert.                                                                                                                   |
| `INVALID_NUMBER_FORMAT` | A format string outside 1 to 255 characters, or a built-in id outside 0 to 49.                                                                                           |
| `INVALID_COLUMN_WIDTH`  | A width outside 0 to 255, or a string other than `"auto"`.                                                                                                               |
| `INVALID_STYLE`         | A colour not `#rrggbb`, a border, alignment, font size or indent outside what Excel accepts.                                                                             |
| `INVALID_FREEZE`        | Frozen rows or columns outside the sheet.                                                                                                                                |
| `INVALID_VIEW`          | A view that is not an object, a zoom that is not a whole number from 10 to 400, `gridlines` or `rightToLeft` that is not a boolean.                                      |
| `INVALID_MERGE`         | A merge not in A1 form, a single cell, outside the sheet, or overlapping another merge.                                                                                  |
| `INVALID_ROW`           | A row that is neither an array of cells nor an object with `cells`.                                                                                                      |
| `INVALID_ROW_HEIGHT`    | A row number below 1, or a height in `rowHeights` or on a row outside 0 to 409 points.                                                                                   |
| `INVALID_FORMULA`       | A formula that is empty, over 8 192 characters as stored or starts with `=`; a cached value that is not text, a number or a boolean.                                     |
| `INVALID_HYPERLINK`     | An `href` that is not an absolute URL or whose scheme runs code, a `cell` not in A1 form or outside the sheet, a `sheet` the workbook lacks.                             |
| `INVALID_VALIDATION`    | A rule without exactly one of `list`, `number` and `date`, a bad range or source, bounds the wrong way round, a message over Excel's lengths.                            |
| `INVALID_PRINT_SETUP`   | An unknown orientation or paper size, a fit, margin or repeat count out of range, header or footer text over 255 characters.                                             |
| `INVALID_IMAGE`         | Bytes that are not a PNG, a width or height that is not positive pixels, a padding that is not pixels from 0, or an image cell that also holds a value, formula or link. |
| `CELL_TEXT_TOO_LONG`    | Text over 32 767 characters.                                                                                                                                             |
| `TOO_MANY_COLUMNS`      | More than 16 384 cells in a row.                                                                                                                                         |
| `TOO_MANY_ROWS`         | More than 1 048 576 rows.                                                                                                                                                |
| `PACKAGE_TOO_LARGE`     | The file, or when streaming one part before compression, would pass 4 GiB; or over 65 529 sheets.                                                                        |
| `ASYNC_ROWS`            | An async iterable of rows given to `writeXlsx`; `streamXlsx` takes those.                                                                                                |
| `NO_DOM`                | `downloadXlsx` called where there is no document.                                                                                                                        |

Empty cells cost nothing: `null`, `undefined` and rows without cells are skipped, and positions stay where you put them.

## Reading (experimental)

`xlsx-mini/read` reads the values of an `.xlsx` file: text, numbers, booleans, dates, the text of an error cell and the result a formula stores. It gives no styles, formulas, comments, pictures or merges, and it never writes to the file it reads. It is built for files from strangers, so a damaged or hostile file is refused with a code, never repaired and never read as an empty workbook. Reading is experimental, and any 0.x minor may change it.

<!-- snippet: read-rows -->

```ts
import { openAsBlob } from "node:fs";
import { openXlsx } from "xlsx-mini/read";

export default async function sumOfTotals(path = "orders.xlsx"): Promise<number> {
  // A Blob is read in slices, never whole.
  const book = await openXlsx(await openAsBlob(path));
  let sum = 0;
  // Each row that holds a value: its number in the sheet from 1, its cells from column A.
  for await (const { number, cells } of book.rows("Orders")) {
    const total = cells[2];
    if (number > 1 && typeof total === "number") sum += total;
  }
  return sum;
}
```

`openXlsx` takes a `Uint8Array`, an `ArrayBuffer` or a `Blob`, such as an uploaded `File`, and reads the package's directory and the sheet list, nothing more. `book.sheets` holds the name of every sheet in workbook order, hidden ones included. `book.rows` takes a name or an index from 0 and reads nothing until you iterate it. Then it yields each row that holds a value as `{ number, cells }`, where `cells` starts at column A, holds `null` for an empty cell and ends at the row's last value. A value is a `string`, a `number`, a `boolean`, a `Date` or `null`. Empty text is `null`, an error cell gives its text, such as `"#DIV/0!"`, and a formula gives the result the file stores, or `null` when the file stores none. Hidden rows and columns read like the others, a merge has its value in its top-left cell, and a chart sheet has no rows. Two sheets can be read at once, and a sheet again.

### Objects by declared columns

`objectsFromRows` is the mirror of `sheetFromObjects`, so one column list writes a sheet of objects and reads it back. The first row is the header, and each later row comes back as `{ number, values }` with the declared keys only.

<!-- snippet: read-objects -->

```ts
import { type ObjectColumn, sheetFromObjects, writeXlsx } from "xlsx-mini";
import { objectsFromRows, openXlsx } from "xlsx-mini/read";

interface Order {
  id: string;
  total: number;
  placed: Date;
}

const columns = [
  { key: "id", header: "Order" },
  { key: "total", header: "Total", style: { numberFormat: '#,##0.00 "kr"' } },
  { key: "placed", header: "Placed", style: { numberFormat: "yyyy-mm-dd" } },
] as const satisfies readonly ObjectColumn<Order>[];

export function exportOrders(orders: Order[]): Promise<Uint8Array> {
  return writeXlsx({ sheets: [sheetFromObjects(orders, { name: "Orders", columns })] });
}

export async function importOrders(file: Uint8Array | Blob): Promise<Order[]> {
  const book = await openXlsx(file);
  const orders: Order[] = [];
  for await (const { number, values } of objectsFromRows(book.rows("Orders"), columns)) {
    const { id, total, placed } = values;
    if (typeof id !== "string" || typeof total !== "number" || !(placed instanceof Date)) {
      throw new Error(`Row ${number} is not an order`);
    }
    orders.push({ id, total, placed });
  }
  return orders;
}
```

Write the column list `as const`, as here, when the list is the declaration: a key it does not declare is then a type error. Type it `ObjectColumn<T>[]` when `T` is the declaration: `values` then has every key of `T` in its type, and a key the list does not declare reads as `undefined`.

### When a file is refused

<!-- snippet: read-errors -->

```ts
import { XlsxReadError, openXlsx } from "xlsx-mini/read";

/** The number of rows in an upload's first sheet, or what to tell the person who sent it. */
export default async function countRows(upload: Blob): Promise<number | string> {
  let count = 0;
  try {
    const book = await openXlsx(upload);
    for await (const _row of book.rows(0)) count++;
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    switch (error.code) {
      case "NOT_XLSX":
        return "This is not an .xlsx file.";
      case "LIMIT_EXCEEDED":
        return `This file asks for more than limits.${error.limit} allows.`;
      case "INVALID_PACKAGE":
      case "INVALID_XML":
      case "INVALID_WORKBOOK":
      case "INVALID_CELL":
        // A damaged file. Another runtime can refuse it with another of these four codes.
        return error.cell === undefined
          ? "This file is damaged."
          : `This file is damaged at cell ${error.cell} of ${error.sheet}.`;
      default:
        // NO_INFLATE, UNKNOWN_SHEET and INVALID_LIMIT are about the runtime or the code;
        // MISSING_HEADER comes from objectsFromRows.
        throw error;
    }
  }
  // A sheet is checked at its last byte, after its rows: trust them once the loop has ended.
  return count;
}
```

Every refusal is an `XlsxReadError` with a `code`, and a workbook whose sheet list or sheet data cannot be found is refused, never read as empty. `rows` yields each row as it reads the sheet, and checks the sheet's size and checksum at its last byte, so act on what you read once the loop has ended without an error. On Bun and Deno, a part whose compressed data is damaged can be refused with `INVALID_PACKAGE` where Node gives `INVALID_XML` or `INVALID_CELL`, so branch on the four codes of a damaged file together, as the snippet does. An error that is not an `XlsxReadError`, such as a failed read of the `Blob`, passes through as it is. The table of error codes at the end of this section lists every code.

### Limits for an upload

<!-- snippet: read-upload -->

```ts
import { XlsxReadError, objectsFromRows, openXlsx } from "xlsx-mini/read";

const columns = [
  { key: "sku", header: "SKU" },
  { key: "counted", header: "Counted" },
] as const;

/** A route handler that takes a stock count of a few thousand rows, uploaded from a form. */
export async function POST(request: Request): Promise<Response> {
  const file = (await request.formData()).get("file");
  if (!(file instanceof Blob)) return new Response("Choose a file.", { status: 400 });
  const counts = new Map<string, number>();
  try {
    // Far below the defaults, which are sized for any file a person uploads.
    const limits = { inflatedBytes: 8 * 1024 * 1024, cellPositions: 50_000 };
    const book = await openXlsx(file, { limits });
    for await (const { number, values } of objectsFromRows(book.rows(0), columns)) {
      if (typeof values.sku !== "string" || typeof values.counted !== "number") {
        return new Response(`Row ${number} needs a SKU and a count.`, { status: 422 });
      }
      counts.set(values.sku, values.counted);
    }
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    return new Response(`The file was refused: ${error.code}.`, { status: 422 });
  }
  return Response.json({ counted: counts.size });
}
```

Every limit has a default sized for a file a person uploads, and you can set any of them lower or higher for one `openXlsx` call. No value turns a limit off. A file past a limit is refused with `LIMIT_EXCEEDED`, and `error.limit` names the limit. The table of limits at the end of this section says what each one bounds.

The limits hold per read. Each iteration of `rows` has its own `inflatedBytes`, `cellPositions` and `rowCharacters`, so two sheets read at once each take up to what one read may, and no limit spans the reads of one workbook. The file decides how many sheets there are, and any number of them may name one part. A file of 6 KB whose 150 sheets name one part, a row of 4 194 304 characters, exhausted a 512 MiB heap when every sheet was read at once. Read the sheets of a file from a stranger one at a time. In trials with the default limits, the most one read reached was 106 MiB of heap on Node, for a row of six cells that each hold the longest text of character references, and 309 MiB resident on Bun, for 30 rows that each hold one cell of 4 194 304 carriage returns. These are measurements, not bounds, and they move with when the engine collects. The opened workbook also holds its shared-string table for as long as you keep the workbook. In trials it was held at up to 2.9 times `sharedStringBytes` on Node and 4.2 times on Bun, for a million entries of 16 characters with one outside Latin-1. `docs/adr/0032-a-values-only-reader-as-a-separate-entry.md` records the design and these trials.

### Where reading runs

On every change, the runtime matrix reads stored and deflated packages in Node, as ESM and as CommonJS, in Bun, in Deno and in Cloudflare workerd. Where the runtime has no `DecompressionStream`, as in the Vercel edge runtime in the matrix, a package of stored parts reads and a deflated part is refused with `NO_INFLATE`. There `writeXlsx` writes stored parts, so its workbooks read back, but a workbook `streamXlsx` wrote does not, because a streamed part is always a deflate stream, even with `compression: "none"`. Deno has no `Blob` backed by a file. Its `fs.openAsBlob` reads the whole file, so on Deno a file opened that way is held in memory whole. No browser is in the matrix.

### What it costs

In a bundle, minified and gzipped, `xlsx-mini/read` alone takes on under 12 kB, and a test holds it to that budget. An application that imports only the writer bundles the same bytes as it would if the package held no read side, which another test checks. Read times and the memory of a read are under Benchmarks.

### Before you trust a value

- A number becomes a `Date` when its cell's number format shows a year, a month or a day, and a date stored as text always does. The `Date`'s UTC fields are the sheet's wall clock, as the writer writes a `Date`, so read it with the UTC getters or `toISOString()`, not in local time. A time of day or an elapsed time stays a number, a fraction of a day, and so does a serial no calendar shows, such as 0 or the 29 February 1900 that never was.
- Text in the format's escape, `_xHHHH_`, reads as the character it names, so `_x0041_` reads as `A`, as Excel and LibreOffice show it. Google Sheets writes such text as it is, without escaping it, so text it showed as `_x0041_` reads as `A` from a file it wrote.

### How it is tested

Every change reads the 87 workbooks in `test/corpus/`, which Excel, Numbers, LibreOffice and Google Sheets wrote, and compares every cell with LibreOffice's reading of the same file, except where `docs/compatibility.md` shows LibreOffice is not the truth. Every change also reads 277 hostile files, each in a process of its own under a time bound and a heap cap, and each must end in the refusal or the values its case states. It reads 5 000 seeded mutations of the corpus too, and each must end in a refusal with a code or a clean read. None may read a sheet that holds a value as a sheet without rows. A scheduled job reads mutations from a fresh seed for 30 minutes every week, and the runtime matrix has each host read what it wrote, as the section on where reading runs says.

### The API of `xlsx-mini/read`

| Export                           | Shape                                                                                                                                                |
| -------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `openXlsx(file, options?)`       | `Promise<OpenedWorkbook>`; `file` is a `Uint8Array`, an `ArrayBuffer` or a `Blob`                                                                    |
| `ReadOptions`                    | `{ limits?: ReadLimits }`                                                                                                                            |
| `OpenedWorkbook`                 | `{ sheets: readonly string[]; rows(sheet: string \| number): AsyncIterable<ReadRow> }`                                                               |
| `ReadRow`, `ReadValue`           | `{ number: number; cells: readonly ReadValue[] }`; `string \| number \| boolean \| Date \| null`                                                     |
| `objectsFromRows(rows, columns)` | `AsyncIterable<ReadObject<Key>>` from any `AsyncIterable<ReadRow>` and a `readonly ReadColumn<Key>[]`                                                |
| `ReadColumn`, `ReadObject`       | `{ key: Key; header?: string }`; `{ number: number; values: { [K in Key]: ReadValue } }`, where `values` has no prototype                            |
| `ReadLimits`                     | `{ inflatedBytes?; cellPositions?; rowCharacters?; sharedStrings?; sharedStringBytes?; parts? }`, each a positive finite number                      |
| `XlsxReadError`, `ReadErrorCode` | `{ code: ReadErrorCode; sheet; cell; limit; header; message }`; each of `sheet`, `cell`, `limit` and `header` is `undefined` where it does not apply |

### Limits

| Limit               | Default   | What it bounds                                                                                                                            |
| ------------------- | --------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `inflatedBytes`     | 128 MiB   | What one operation inflates. Opening the workbook, with its shared strings and styles, is one operation, and so is each `rows` iteration. |
| `cellPositions`     | 5 000 000 | The positions one `rows` iteration yields, empty cells before a value included.                                                           |
| `rowCharacters`     | 4 194 304 | The text one row holds. A row is held whole, so this bounds the text of the values one read holds at a time.                              |
| `sharedStrings`     | 1 000 000 | Entries in the shared-string table.                                                                                                       |
| `sharedStringBytes` | 32 MiB    | What the shared-string table inflates to.                                                                                                 |
| `parts`             | 100 000   | The parts in the package, read or not. A workbook with one picture per row has about one part per row.                                    |

A part's declared size is charged before a byte of it is inflated. The parts that become lists in memory and the package's directory share an eighth of `inflatedBytes`. Three rules are fixed and are not limits: elements nest at most 256 deep, one token holds at most 4 194 304 characters, and the names of the open elements and the namespace declarations in scope hold as many together, each declaration counting 4 096 more than its prefix and its namespace name. A part that breaks one is refused with `INVALID_XML`.

### Error codes

A refusal while `rows` reads a sheet names the sheet in `sheet`, except a refusal of the shared strings or the styles, which belong to the whole workbook.

| Code               | Means                                                                                                                                                                                                                                | Carries                                             |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------- |
| `NOT_XLSX`         | Not a package, or a package that is not a workbook, such as an old `.xls`, an encrypted workbook, a CSV or another kind of zip.                                                                                                      | nothing                                             |
| `INVALID_PACKAGE`  | The package is damaged or contradicts itself, for example in its directory, a local header, two parts with one name, a part's size or CRC-32, or a compression method.                                                               | `sheet` while `rows` reads one                      |
| `INVALID_XML`      | A part is not well-formed XML in UTF-8, holds a DOCTYPE, or breaks one of the fixed rules above.                                                                                                                                     | `sheet` while `rows` reads one                      |
| `INVALID_WORKBOOK` | The parts do not make a workbook: the sheet list, a sheet's data, the shared strings or the styles are missing or damaged, a relationship leaves the package, two sheets share a name, or the date system is neither true nor false. | `sheet` when the refusal is about one sheet         |
| `INVALID_CELL`     | A row or a cell is out of order, a reference is malformed, a value is not what its cell's type says, or a shared string or a style does not exist.                                                                                   | `sheet`, and `cell` where the place has a reference |
| `LIMIT_EXCEEDED`   | The file asks for more than a limit allows.                                                                                                                                                                                          | `limit`, and `sheet` while `rows` reads one         |
| `MISSING_HEADER`   | `objectsFromRows` found no header row, or one that lacks a declared column.                                                                                                                                                          | `header`, the caller's own text                     |
| `NO_INFLATE`       | A part is deflated, and the runtime has no `DecompressionStream`, or one that cannot be held to the end of a deflate stream.                                                                                                         | `sheet` while `rows` reads one                      |
| `UNKNOWN_SHEET`    | `rows` was called with a name or an index the workbook does not have. It throws at the call.                                                                                                                                         | `sheet` when called with a name                     |
| `INVALID_LIMIT`    | A limit that is not a positive finite number.                                                                                                                                                                                        | `limit`                                             |

## API

| Export                              | Shape                                                                                                                                                                                                                                                                          |
| ----------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `writeXlsx(workbook, options?)`     | `Promise<Uint8Array<ArrayBuffer>>`                                                                                                                                                                                                                                             |
| `WriteOptions`                      | `{ coerce?: Coerce; compression?: "deflate" \| "none" }`; `"none"` writes STORE entries, byte-identical on every platform                                                                                                                                                      |
| `streamXlsx(workbook, options?)`    | `ReadableStream<Uint8Array>`; the workbook may be a `StreamedWorkbook`                                                                                                                                                                                                         |
| `StreamOptions`                     | `WriteOptions` plus `strings?: "inline" \| "shared"`; inline is the default and keeps memory flat                                                                                                                                                                              |
| `Workbook`                          | `{ sheets: readonly Sheet[] }`                                                                                                                                                                                                                                                 |
| `StreamedWorkbook`, `StreamedSheet` | `{ sheets: readonly StreamedSheet[] }`; a `Sheet` whose `rows` may also be an `AsyncIterable<Row>`                                                                                                                                                                             |
| `Sheet`                             | `{ name; rows: Iterable<Row>; columns?: (Column \| undefined)[]; freeze?: Freeze; view?: SheetView; autoFilter?: boolean; autoWidth?: boolean; merges?: string[]; rowHeights?: Record<number, number>; validations?: DataValidation[]; print?: PrintSetup; hidden?: boolean }` |
| `Row`, `Cell`                       | `readonly Cell[] \| RowObject`; `CellValue \| CellObject \| FormulaCell \| LinkCell \| SheetLinkCell \| ImageCell`                                                                                                                                                             |
| `RowObject`                         | `{ cells: readonly Cell[]; style?: CellStyle; height?: number; hidden?: boolean }`; the style reaches every position in `cells`, the height is in points                                                                                                                       |
| `CellObject`, `FormulaCell`         | `{ value: CellValue; style? }`; `{ formula: string; value?: CellValue; style? }`                                                                                                                                                                                               |
| `LinkCell`, `SheetLinkCell`         | `{ value; href: string; style? }`; `{ value; sheet?: string; cell: string; style? }`                                                                                                                                                                                           |
| `ImageCell`                         | `{ image: Uint8Array; width?: number; height?: number; padding?: number; style? }`; a PNG anchored to the cell, sizes and the space around it in pixels                                                                                                                        |
| `CellValue`                         | `string \| number \| boolean \| Date \| TemporalLike \| null \| undefined \| Registered`                                                                                                                                                                                       |
| `CellStyle`                         | `{ bold?; italic?; underline?; fontSize?; fontColor?; fill?; border?; align?; verticalAlign?; wrap?; indent?; numberFormat? }`; colours `#rrggbb`                                                                                                                              |
| `DataValidation`                    | `{ range; allowBlank?; input?; error? }` with one of `list: string[] \| string`, `number: { min?; max?; integer? }`, `date: { min?: Date; max?: Date }`                                                                                                                        |
| `PrintSetup`                        | `{ orientation?; paperSize?: PaperSize; fitTo?: { width?; height? }; margins?: PrintMargins; header?; footer?; repeatRows? }`                                                                                                                                                  |
| `BorderStyle`, `Borders`            | `"thin" \| "medium" \| "thick"`; `{ top?; right?; bottom?; left? }` of those                                                                                                                                                                                                   |
| `HorizontalAlign`, `VerticalAlign`  | `"left" \| "center" \| "right"`; `"top" \| "middle" \| "bottom"`                                                                                                                                                                                                               |
| `Column`, `Freeze`                  | `{ width?: number \| "auto"; style?: CellStyle; hidden?: boolean }`; `{ rows?: number; columns?: number }`                                                                                                                                                                     |
| `SheetView`                         | `{ gridlines?: boolean; zoom?: number; rightToLeft?: boolean }`; how the sheet is shown when it opens, the zoom in percent from 10 to 400                                                                                                                                      |
| `Coerce`, `Register`, `Registered`  | `(value: Date \| TemporalLike \| Registered, path: string) => CellValue`; the augmentation point and what it holds                                                                                                                                                             |
| `sheetFromObjects(data, options)`   | `Sheet`; options `{ name; columns?: ObjectColumn<T>[]; header?: CellStyle \| false; rowStyle?: (item: T, index: number) => CellStyle \| undefined; freeze?: boolean; autoFilter?: boolean; autoWidth?: boolean }`                                                              |
| `XlsxMiniError`, `ErrorCode`        | `{ code: ErrorCode; path: string; message }`                                                                                                                                                                                                                                   |
| `XLSX_MIME_TYPE`                    | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`                                                                                                                                                                                                            |
| `downloadXlsx(bytes, filename)`     | from `xlsx-mini/browser`; adds `.xlsx` when the name lacks it                                                                                                                                                                                                                  |

Rows are read once: pass a generator or a database cursor and the writer streams through it, encoding each row as it comes.

Below 1.0.0 a minor version may change this API; 1.0.0 follows once it has survived real use unchanged, with no code change of its own.

## What it does not do

Four things are outside this library on purpose. Each is large against a writer of this size, and leaving them out is what keeps it small:

- **Editing a workbook.** `xlsx-mini/read` reads the values of a workbook and nothing else, so a workbook cannot be opened, changed and saved. Read its values and write a new workbook with `writeXlsx`. Styles, formulas, pictures and everything else that is not a value stay behind.
- **Charts.** A chart is a drawing format of its own, larger than everything here together. Write the data; the chart is made in the application.
- **Pivot tables.** A pivot table needs a cache of its source data in a second format beside the sheet.
- **Password encryption.** An encrypted workbook is a different container around the package, with cryptography of its own. Encrypt the bytes with a tool made for that. `xlsx-mini/read` refuses an encrypted workbook with `NOT_XLSX`.

`docs/adr/0026-non-goals.md` records the reasons and what would reopen each.

## Security

No runtime dependencies and no install scripts. The writer touches no file, network or global state; it reads what you give it and returns bytes. Text is escaped for XML and characters XML cannot carry are dropped, so cell content cannot break the file. Text never becomes a formula, whatever it starts with; a formula or link is only ever a plain object you build with `formula`, `href` or `cell`, so convert untrusted objects to text before putting them in a cell. Link targets whose scheme runs code (`javascript:`, `vbscript:`, `data:`) are refused. The read side opens no file and fetches nothing. It reads the bytes or the `Blob` you give it, refuses a damaged file with a code instead of repairing it, holds each read to its limits, makes no property name from text in the file, and resolves nothing outside the package. Releases are published through npm trusted publishing with provenance, so every version on npm traces to a commit and a workflow run in this repository. To report a vulnerability, see `SECURITY.md`.

## Benchmarks

Every pull request gets one comment: the pull request and `main` built and timed on the same runner, in turns, on Node, Bun and Deno, for three suites. `pnpm bench` runs them here. On an Apple M5 Pro with Node 24, the minimum of sixteen runs, each benchmark in processes of its own:

| Suite                                        | Deflate |  STORE | Streamed |
| -------------------------------------------- | ------: | -----: | -------: |
| Wide table, 100 000 rows by 10 mixed columns |  463 ms | 141 ms |   447 ms |
| Many sheets, 40 sheets of 2 500 rows by 6    |  232 ms |  68 ms |   197 ms |
| String-heavy, 200 000 rows by 5 text columns |  370 ms |  89 ms |   698 ms |

STORE is the writer's own path; deflate adds the runtime's zlib and is the default. Streamed is `streamXlsx` with its defaults into a sink that discards the bytes: inline strings cost more where text repeats, which is what `strings: "shared"` is for. Streaming the wide table at a million rows holds about 10 MiB of heap, the same as at a hundred thousand. The comment also shows what the pull request does to the size of the package in a bundle. Budgets and the method are in `docs/adr/0021-performance-baseline-and-budgets.md`.

Reading is timed the same way, on the file the writer makes of each suite by default. On the same machine with Node 24, the minimum of sixteen reads of every row of every sheet, each benchmark in processes of its own:

| Suite, what one read gives                                       |   Read |
| ---------------------------------------------------------------- | -----: |
| Wide table, 100 001 rows of 10 cells                             | 314 ms |
| Wide table through `objectsFromRows`, 100 000 objects of 10 keys | 342 ms |
| Many sheets, 40 sheets of 2 501 rows of 6 cells                  | 191 ms |
| String-heavy, 200 001 rows of 5 cells                            | 288 ms |

A read holds about as much at a million rows as at 100 000. Read from bytes on Node 24, the wide table streamed with inline strings held at most 1.0 MiB of heap and array buffers above what the process held before the read at 100 000 rows, and 1.3 MiB at a million, with `inflatedBytes` and `cellPositions` raised because a million of its rows pass both defaults.

## Compatibility

Each golden fixture is opened by hand in four applications before a release; `docs/compatibility.md` records how. For the current fixtures:

| Fixture                          | Excel 16.114 (macOS) | Numbers 14.4 | Google Sheets | LibreOffice 26.8 |
| -------------------------------- | :------------------: | :----------: | :-----------: | :--------------: |
| Strings and numbers              |        clean         |    clean     |     clean     |      clean       |
| Every kind of cell value         |        clean         |    clean     |     clean     |      clean       |
| Deflated entries                 |        clean         |    clean     |     clean     |      clean       |
| Several sheets, escaped names    |        clean         |    clean     |     clean     |      clean       |
| Styles and number formats        |        clean         |    clean     |     clean     |      clean       |
| Widths, frozen panes, autofilter |        clean         |    clean     |     clean     |      clean       |
| Built from objects               |        clean         |    clean     |     clean     |      clean       |
| Fonts, fills, borders, merges    |        clean         |    clean     |     clean     |      clean       |
| Streamed, descriptors, inline    |        clean         |    clean     |     clean     |      clean       |
| Formulas, cached and not         |        clean         |    clean     |     clean     |      clean       |
| Functions newer than Excel 2007  |        clean         |    clean     |    8 of 10    |      clean       |
| Hyperlinks, external and within  |        clean         |    clean     |     clean     |      clean       |
| Data validation                  |        clean         |    clean     |     clean     |      clean       |
| Auto column widths               |        clean         |    clean     |     clean     |      clean       |
| Image cells                      |        clean         |    clean     |     clean     |      clean       |
| Print setup                      |        clean         |    clean     |     clean     |      clean       |
| Row styles                       |        clean         |    clean     |     clean     |      clean       |
| Hidden rows, columns and sheets  |        clean         | sheet shown  |     clean     |      clean       |
| Gridlines, zoom, right-to-left   |        clean         |   no zoom    |    no zoom    |      clean       |
| Image padding                    |        clean         |    clean     |     clean     |      clean       |
| Text that looks like an escape   |        clean         |    clean     |     clean     |     13 of 14     |

Google Sheets computes eight of the ten formulas in the newer-functions fixture: it does not know `FILTER` under the name Excel stores it by, and its own `CONCAT` takes two values where Excel's takes a range. Numbers has no hidden sheets: it shows a hidden sheet as an ordinary one and opens on the first sheet, and it keeps a picture in a hidden row or column in view. Hidden rows and columns it hides. Neither Numbers nor Google Sheets takes a sheet's zoom from the file: both open at their own. Text that looks like the format's `_xHHHH_` escape reads back as written. In a text that also holds a line feed, LibreOffice reads every carriage return as a line break. Numbers shows the escape in inline strings, which `streamXlsx` writes unless given `strings: "shared"`. LibreOffice also reads every fixture back cell for cell in CI, on every change, formula results included.

## Contributing

Issues and pull requests are welcome; `CONTRIBUTING.md` is short. The design decisions live in `docs/adr/`.

## License

MIT
