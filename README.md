# xlsx-mini

[![npm](https://img.shields.io/npm/v/xlsx-mini)](https://www.npmjs.com/package/xlsx-mini)
[![ci](https://github.com/ANicholasson/xlsx-mini/actions/workflows/ci.yml/badge.svg)](https://github.com/ANicholasson/xlsx-mini/actions/workflows/ci.yml)
[![license](https://img.shields.io/npm/l/xlsx-mini)](./LICENSE)

Writes `.xlsx` files from TypeScript with no dependencies, on every JavaScript runtime. You describe sheets, rows and cells; it returns the bytes. The files open without a repair dialog in Excel, Numbers, Google Sheets and LibreOffice, and the writer is measured on every change. Every TypeScript snippet in this README is run by the test suite.

## Install

```sh
pnpm add xlsx-mini
```

ESM only. Node 20.12 or later (22 or later for `require`), Bun, Deno, Cloudflare Workers, Vercel's edge runtime and evergreen browsers: the core uses only web-standard APIs, and the runtime matrix in CI proves each host on every change.

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

`streamXlsx` returns a web `ReadableStream` of the file's bytes, produced as the consumer reads them, and takes rows from an async iterable, so a database cursor can feed a sheet page by page. Peak memory stays flat however many rows go out. Hand the stream to a `Response`, pipe it to a file, or `for await` over its chunks.

<!-- snippet: streaming -->

```ts
// app/api/orders/route.ts
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

export function GET(fetchPage: (page: number) => Promise<Row[]>): Response {
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

Text is written into each cell rather than through a shared-strings table, which is what keeps memory flat; `strings: "shared"` keeps the table instead, for a smaller file at the cost of holding every distinct string until the end. The stream throws before it exists when the workbook itself is invalid, and errors mid-way when a cell cannot be written. Cancelling it, as a client going away does, closes the row source.

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

A style also sets the font (`italic`, `fontSize` in points, `fontColor`), a solid `fill`, a `border` for all four sides or per side (`thin`, `medium`, `thick`), `align`, `verticalAlign`, `wrap` and `indent`. An indent aligns left unless `align` is `right`. Colours are `#rrggbb`. A sheet merges ranges with `merges` and sets `rowHeights` in points by row number. A cell object with a style and a `null` value writes an empty styled cell, which is how a border or a fill runs across a blank.

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

| Code                    | When                                                                                           |
| ----------------------- | ---------------------------------------------------------------------------------------------- |
| `NO_SHEETS`             | The workbook has no sheet.                                                                     |
| `INVALID_SHEET_NAME`    | Empty, over 31 characters, contains `[ ] : * ? / \`, starts or ends with `'`, or is `History`. |
| `DUPLICATE_SHEET_NAME`  | Two sheets share a name, ignoring case.                                                        |
| `INVALID_NUMBER`        | `NaN` or an infinity.                                                                          |
| `INVALID_DATE`          | An invalid `Date`.                                                                             |
| `DATE_OUT_OF_RANGE`     | Before 1900-01-01 or after 9999-12-31.                                                         |
| `UNSUPPORTED_VALUE`     | A value no cell can hold and `coerce` did not convert.                                         |
| `INVALID_NUMBER_FORMAT` | A format string outside 1 to 255 characters, or a built-in id outside 0 to 49.                 |
| `INVALID_COLUMN_WIDTH`  | A width outside 0 to 255.                                                                      |
| `INVALID_STYLE`         | A colour not `#rrggbb`, a border, alignment, font size or indent outside what Excel accepts.   |
| `INVALID_FREEZE`        | Frozen rows or columns outside the sheet.                                                      |
| `INVALID_MERGE`         | A merge not in A1 form, a single cell, outside the sheet, or overlapping another merge.        |
| `INVALID_ROW_HEIGHT`    | A row number below 1 or a height outside 0 to 409 points.                                      |
| `CELL_TEXT_TOO_LONG`    | Text over 32 767 characters.                                                                   |
| `TOO_MANY_COLUMNS`      | More than 16 384 cells in a row.                                                               |
| `TOO_MANY_ROWS`         | More than 1 048 576 rows.                                                                      |
| `PACKAGE_TOO_LARGE`     | The file would pass 4 GiB.                                                                     |
| `ASYNC_ROWS`            | An async iterable of rows given to `writeXlsx`; `streamXlsx` takes those.                      |
| `NO_DOM`                | `downloadXlsx` called where there is no document.                                              |

Empty cells cost nothing: `null`, `undefined` and rows without cells are skipped, and positions stay where you put them.

## API

| Export                              | Shape                                                                                                                                                             |
| ----------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `writeXlsx(workbook, options?)`     | `Promise<Uint8Array<ArrayBuffer>>`                                                                                                                                |
| `WriteOptions`                      | `{ coerce?: Coerce; compression?: "deflate" \| "none" }`; `"none"` writes STORE entries, byte-identical on every platform                                         |
| `streamXlsx(workbook, options?)`    | `ReadableStream<Uint8Array>`; the workbook may be a `StreamedWorkbook`                                                                                            |
| `StreamOptions`                     | `WriteOptions` plus `strings?: "inline" \| "shared"`; inline is the default and keeps memory flat                                                                 |
| `Workbook`                          | `{ sheets: readonly Sheet[] }`                                                                                                                                    |
| `StreamedWorkbook`, `StreamedSheet` | `{ sheets: readonly StreamedSheet[] }`; a `Sheet` whose `rows` may also be an `AsyncIterable<Row>`                                                                |
| `Sheet`                             | `{ name; rows: Iterable<Row>; columns?: (Column \| undefined)[]; freeze?: Freeze; autoFilter?: boolean; merges?: string[]; rowHeights?: Record<number, number> }` |
| `Row`, `Cell`                       | `readonly Cell[]`; `CellValue \| { value: CellValue; style?: CellStyle }`                                                                                         |
| `CellValue`                         | `string \| number \| boolean \| Date \| TemporalLike \| null \| undefined \| Registered`                                                                          |
| `CellStyle`                         | `{ bold?; italic?; fontSize?; fontColor?; fill?; border?; align?; verticalAlign?; wrap?; indent?; numberFormat? }`; colours `#rrggbb`                             |
| `BorderStyle`, `Borders`            | `"thin" \| "medium" \| "thick"`; `{ top?; right?; bottom?; left? }` of those                                                                                      |
| `HorizontalAlign`, `VerticalAlign`  | `"left" \| "center" \| "right"`; `"top" \| "middle" \| "bottom"`                                                                                                  |
| `Column`, `Freeze`                  | `{ width?: number; style?: CellStyle }`; `{ rows?: number; columns?: number }`                                                                                    |
| `Coerce`, `Register`, `Registered`  | `(value: Date \| TemporalLike \| Registered, path: string) => CellValue`; the augmentation point and what it holds                                                |
| `sheetFromObjects(data, options)`   | `Sheet`; options `{ name; columns?: ObjectColumn<T>[]; header?: CellStyle \| false; freeze?: boolean; autoFilter?: boolean }`                                     |
| `XlsxMiniError`, `ErrorCode`        | `{ code: ErrorCode; path: string; message }`                                                                                                                      |
| `XLSX_MIME_TYPE`                    | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`                                                                                               |
| `downloadXlsx(bytes, filename)`     | from `xlsx-mini/browser`; adds `.xlsx` when the name lacks it                                                                                                     |

Rows are read once: pass a generator or a database cursor and the writer streams through it, encoding each row as it comes.

Below 1.0.0 a minor version may change this API; 1.0.0 follows once it has survived real use unchanged, with no code change of its own.

## Security

No runtime dependencies and no install scripts. The writer touches no file, network or global state; it reads what you give it and returns bytes. Text is escaped for XML and characters XML cannot carry are dropped, so cell content cannot break the file. Releases are published through npm trusted publishing with provenance, so every version on npm traces to a commit and a workflow run in this repository. To report a vulnerability, see `SECURITY.md`.

## Benchmarks

Every pull request gets one comment: the pull request and `main` built and timed on the same runner, in turns, on Node, Bun and Deno, for three suites. `pnpm bench` runs them here. On an Apple M5 Pro with Node 24, the minimum of seven runs:

| Suite                                        | Deflate |  STORE | Streamed |
| -------------------------------------------- | ------: | -----: | -------: |
| Wide table, 100 000 rows by 10 mixed columns |  495 ms | 177 ms |   522 ms |
| Many sheets, 40 sheets of 2 500 rows by 6    |  238 ms |  74 ms |   248 ms |
| String-heavy, 200 000 rows by 5 text columns |  367 ms |  90 ms |   913 ms |

STORE is the writer's own path; deflate adds the runtime's zlib and is the default. Streamed is `streamXlsx` with its defaults into a sink that discards the bytes: inline strings cost more where text repeats, which is what `strings: "shared"` is for. Streaming the wide table at a million rows holds about 10 MiB of heap, the same as at a hundred thousand. Budgets and the method are in `docs/adr/0021-performance-baseline-and-budgets.md`.

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

LibreOffice also reads every fixture back cell for cell in CI, on every change.

## Contributing

Issues and pull requests are welcome; `CONTRIBUTING.md` is short. The design decisions live in `docs/adr/`.

## License

MIT
