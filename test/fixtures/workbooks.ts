import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { type WriteOptions, type Workbook, sheetFromObjects } from "../../src/index.js";

const imagesDir = join(dirname(fileURLToPath(import.meta.url)), "images");
/** A 120 by 40 picture of vertical bars and a 48 by 48 badge, both PNG, both committed. */
export const stripesPng = new Uint8Array(readFileSync(join(imagesDir, "stripes.png")));
export const badgePng = new Uint8Array(readFileSync(join(imagesDir, "badge.png")));

/** The workbook behind minimal.xlsx. Change it only together with the fixture. */
export const minimalWorkbook: Workbook = {
  sheets: [
    {
      name: "Sheet1",
      rows: [
        ["Name", "City", "Amount"],
        ["Kari Nordmann", "Oslo", 1234.5],
        ["Ola <Nordmann> & co", "Tromsø", -42],
        ["Ærlig «ekte» 🎉", null, 1e21],
        [" padded ", "", 0],
      ],
    },
  ],
};

const when = Date.UTC(2026, 10, 4, 13, 45, 30);
const wallClock = {
  year: 2026,
  month: 11,
  day: 4,
  hour: 13,
  minute: 45,
  second: 30,
  millisecond: 0,
  microsecond: 0,
  nanosecond: 0,
} as const;

/** The workbook behind cell-types.xlsx: one row per kind of value the writer accepts. */
export const cellTypesWorkbook: Workbook = {
  sheets: [
    {
      name: "Cell types",
      rows: [
        ["Kind", "Value"],
        ["text", "plain text"],
        ["number", 3.14],
        ["boolean true", true],
        ["boolean false", false],
        ["Date (UTC 2026-11-04 13:45:30)", new Date(when)],
        [
          "Temporal.PlainDate",
          { [Symbol.toStringTag]: "Temporal.PlainDate", year: 2026, month: 11, day: 4 },
        ],
        [
          "Temporal.PlainDateTime",
          { [Symbol.toStringTag]: "Temporal.PlainDateTime", ...wallClock },
        ],
        [
          "Temporal.ZonedDateTime",
          { [Symbol.toStringTag]: "Temporal.ZonedDateTime", ...wallClock },
        ],
        ["Temporal.Instant", { [Symbol.toStringTag]: "Temporal.Instant", epochMilliseconds: when }],
        ["wrapped { value }", { value: 42 }],
        ["empty", null],
        ["control characters stripped", "a\u0000b\u0007c"],
        ["newline kept", "line1\nline2"],
        ["bigint through coerce", 12345678901234n],
      ],
    },
  ],
};

export const cellTypesOptions: WriteOptions = {
  coerce: (value) => (typeof value === "bigint" ? Number(value) : value),
};

/** The workbook behind multi-sheet.xlsx: three sheets, names that need escaping, text shared across sheets. */
export const multiSheetWorkbook: Workbook = {
  sheets: [
    {
      name: "Orders & Returns",
      rows: [
        ["Order", "Status"],
        ["A-1001", "shipped"],
        ["A-1002", "returned"],
      ],
    },
    {
      name: "Ærlig <oversikt>",
      rows: [
        ["Status", "Count"],
        ["shipped", 1],
        ["returned", 1],
      ],
    },
    { name: "Empty", rows: [] },
  ],
};

const kroner = { numberFormat: '#,##0.00 "kr"' };
const header = { bold: true };
const day = (y: number, m: number, d: number) => new Date(Date.UTC(y, m - 1, d));

/** The workbook behind styles.xlsx: bold header, currency and percent columns, ISO and default dates, a bold total row. */
export const stylesWorkbook: Workbook = {
  sheets: [
    {
      name: "Orders",
      columns: [
        undefined,
        { style: kroner },
        { style: { numberFormat: "0.0%" } },
        { style: { numberFormat: "yyyy-mm-dd" } },
      ],
      rows: [
        [
          { value: "Order", style: header },
          { value: "Total", style: header },
          { value: "Discount", style: header },
          { value: "Ordered", style: header },
          { value: "Delivered", style: header },
        ],
        ["A-1001", 1234.5, 0.1, day(2026, 11, 1), new Date(Date.UTC(2026, 10, 3, 14, 30))],
        ["A-1002", 99, 0, day(2026, 11, 2), new Date(Date.UTC(2026, 10, 4, 9, 15))],
        ["A-1003", 100000, 0.255, day(2026, 11, 3), null],
        [
          { value: "Sum", style: header },
          { value: 101333.5, style: header },
        ],
      ],
    },
  ],
};

/** The workbook behind layout.xlsx: column widths, a frozen header row and first column, an autofilter. */
export const layoutWorkbook: Workbook = {
  sheets: [
    {
      name: "Orders 2026",
      columns: [{ width: 12 }, { width: 32 }, { width: 14, style: kroner }],
      freeze: { rows: 1, columns: 1 },
      autoFilter: true,
      rows: [
        [
          { value: "Order", style: header },
          { value: "Customer", style: header },
          { value: "Total", style: header },
        ],
        ["A-1001", "Kari Nordmann", 1234.5],
        ["A-1002", "Ola Nordmann", 99],
        ["A-1003", "Per O'Brien", 100000],
      ],
    },
    { name: "Notes", rows: [["No filter here"]] },
  ],
};

/** The workbook behind orders.xlsx: the README example for sheetFromObjects. */
export const ordersWorkbook: Workbook = {
  sheets: [
    sheetFromObjects(
      [
        {
          id: "A-1001",
          customer: "Kari Nordmann",
          total: 1234.5,
          placedAt: new Date(Date.UTC(2026, 10, 1, 9, 30)),
        },
        {
          id: "A-1002",
          customer: "Ola Nordmann",
          total: 99,
          placedAt: new Date(Date.UTC(2026, 10, 2, 14, 0)),
        },
        {
          id: "A-1003",
          customer: "Per O'Brien",
          total: 100000,
          placedAt: new Date(Date.UTC(2026, 10, 3, 8, 15)),
        },
      ],
      {
        name: "Orders",
        columns: [
          { key: "id", header: "Order", width: 12 },
          { key: "customer", header: "Customer", width: 24 },
          { key: "total", header: "Total", width: 14, style: { numberFormat: '#,##0.00 "kr"' } },
          {
            key: "placedAt",
            header: "Placed",
            width: 18,
            style: { numberFormat: "yyyy-mm-dd hh:mm" },
          },
        ],
      },
    ),
  ],
};

const title = {
  bold: true,
  fontSize: 16,
  fontColor: "#ffffff",
  fill: "#1f4e79",
  align: "center",
  verticalAlign: "middle",
} as const;
const heading = { bold: true, fill: "#d9e1f2", border: "thin", verticalAlign: "middle" } as const;
const ruled = { top: "medium", bottom: "thick" } as const;

/** The workbook behind appearance.xlsx: a merged title, a filled and bordered header, wrapped and indented text, row heights, a ruled total. */
export const appearanceWorkbook: Workbook = {
  sheets: [
    {
      name: "Invoice",
      columns: [
        { width: 14 },
        { width: 36 },
        { width: 8, style: { align: "right" } },
        { width: 14, style: kroner },
      ],
      merges: ["A1:D1"],
      rowHeights: { 1: 28, 3: 48 },
      rows: [
        [{ value: "Invoice 2026-1042", style: title }],
        [
          { value: "Item", style: heading },
          { value: "Description", style: heading },
          { value: "Qty", style: { ...heading, align: "right" } },
          { value: "Amount", style: { ...heading, align: "right" } },
        ],
        [
          "Desk",
          {
            value: "Height-adjustable desk, oak veneer, delivered and assembled on site",
            style: { wrap: true, verticalAlign: "top" },
          },
          2,
          9998,
        ],
        ["Chair", { value: "Mesh office chair", style: { indent: 1 } }, 4, 5198],
        [
          { value: "Total", style: { bold: true, border: ruled } },
          { value: null, style: { border: ruled } },
          { value: 6, style: { bold: true, border: ruled } },
          { value: 15196, style: { bold: true, border: ruled } },
        ],
        [],
        [{ value: "Payment within 30 days.", style: { italic: true, fontColor: "#808080" } }],
      ],
    },
  ],
};

const money = { numberFormat: '#,##0.00 "kr"' } as const;

/**
 * The workbook behind future-functions.xlsx: functions newer than Excel 2007 as a caller types
 * them, one already under its prefix, a function's name inside quoted text, and an older function.
 */
export const futureFunctionsWorkbook: Workbook = {
  sheets: [
    {
      name: "Functions",
      columns: [{ width: 18 }, { width: 22 }, { width: 4 }, { width: 6 }, { width: 6 }],
      rows: [
        [
          { value: "Function", style: header },
          { value: "Result", style: header },
          null,
          { value: "n", style: header },
          { value: "t", style: header },
        ],
        ["STDEV.S", { formula: "ROUND(STDEV.S(D2:D9),4)" }, null, 2, "a"],
        ["IFS", { formula: 'IFS(D2>5,"big",D2>1,"mid",TRUE,"small")' }, null, 4, "b"],
        ["TEXTJOIN", { formula: 'TEXTJOIN("-",TRUE,E2:E4)' }, null, 4, "c"],
        ["CONCAT", { formula: "CONCAT(E2:E4)" }, null, 4],
        ["XLOOKUP", { formula: 'XLOOKUP("b",E2:E4,D2:D4)' }, null, 5],
        ["MAXIFS", { formula: 'MAXIFS(D2:D9,D2:D9,"<6")' }, null, 5],
        ["FILTER in a sum", { formula: "SUM(FILTER(D2:D9,D2:D9>4))" }, null, 7],
        ["Already prefixed", { formula: "_xlfn.XLOOKUP(7,D2:D9,D2:D9)" }, null, 9],
        ["Name in a text", { formula: 'IF(D2=2,"IFS(x) and XLOOKUP(","no")' }],
        ["Older function", { formula: "SUM(D2:D9)" }],
      ],
    },
  ],
};

const reportHeader = { bold: true, fill: "#1f2937", fontColor: "#ffffff" } as const;
const band = { fill: "#f3f4f6" } as const;
const totalRow = { bold: true, border: { top: "thin" } } as const;

/**
 * The workbook behind row-styles.xlsx: a header styled by its row, with a height; banded rows
 * between rows given as arrays; a cell overriding one property of its row; a styled row with an
 * empty cell; a total row with a border; and row styles over a column's number format.
 */
export const rowStylesWorkbook: Workbook = {
  sheets: [
    {
      name: "Report",
      columns: [{ width: 14 }, { width: 16 }, { width: 10 }, { width: 16, style: money }],
      rows: [
        { cells: ["Order", "Customer", "Items", "Total"], style: reportHeader, height: 24 },
        ["A-1001", "Kari", 3, 1234.5],
        { cells: ["A-1002", "Ola", 1, 99], style: band },
        ["A-1003", "Liv", 12, 560.25],
        { cells: ["A-1004", null, 2, { value: 75, style: { fill: "#fde68a" } }], style: band },
        {
          cells: ["Total", null, { formula: "SUM(C2:C5)" }, { formula: "SUM(D2:D5)" }],
          style: totalRow,
        },
      ],
    },
  ],
};

/** The workbook behind formulas.xlsx: per-row products, a sum, cached results, and text that only looks like a formula. */
export const formulasWorkbook: Workbook = {
  sheets: [
    {
      name: "Budget",
      columns: [
        { width: 16 },
        { width: 8 },
        { width: 12, style: money },
        { width: 14, style: money },
      ],
      rows: [
        [
          { value: "Item", style: header },
          { value: "Qty", style: header },
          { value: "Price", style: header },
          { value: "Total", style: header },
        ],
        ["Desk", 2, 4999, { formula: "B2*C2" }],
        ["Chair", 4, 1299.5, { formula: "B3*C3" }],
        ["Lamp", 1, 349, { formula: "B4*C4" }],
        [
          { value: "Sum", style: { bold: true } },
          { formula: "SUM(B2:B4)", value: 7, style: { bold: true } },
          null,
          { formula: "SUM(D2:D4)", style: { bold: true } },
        ],
        [],
        ["Items", { formula: "COUNTA(A2:A4)", value: 3 }],
        ["Over budget", { formula: "D5>10000", value: true }],
        ["Label", { formula: 'CONCATENATE("Items: ",COUNTA(A2:A4))', value: "Items: 3" }],
        ["Not a formula", "=SUM(B2:B4)"],
      ],
    },
  ],
};

const link = { fontColor: "#0563c1", underline: true } as const;

/** The workbook behind hyperlinks.xlsx: external targets, a mail address, and jumps within the workbook. */
export const hyperlinksWorkbook: Workbook = {
  sheets: [
    {
      name: "Links",
      columns: [{ width: 14 }, { width: 28 }],
      rows: [
        [
          { value: "Kind", style: header },
          { value: "Link", style: header },
        ],
        ["Web", { value: "Example docs", href: "https://example.com/docs?a=1&b=2", style: link }],
        ["Mail", { value: "Write to us", href: "mailto:hello@example.com", style: link }],
        ["Sheet", { value: "See the summary", sheet: "Summary", cell: "B2", style: link }],
        ["Here", { value: "Back to the top", cell: "A1", style: link }],
      ],
    },
    {
      name: "Summary",
      rows: [
        ["Total", 42],
        [{ value: "Back to the links", sheet: "Links", cell: "A1:B5", style: link }],
      ],
    },
  ],
};

/** The workbook behind validation.xlsx: dropdowns from a list and from a range, number and date bounds, messages. */
export const validationWorkbook: Workbook = {
  sheets: [
    {
      name: "Orders",
      columns: [
        { width: 12 },
        { width: 14 },
        { width: 12 },
        { width: 10 },
        { width: 14 },
        { width: 10 },
      ],
      rows: [
        [
          { value: "Order", style: header },
          { value: "Status", style: header },
          { value: "Priority", style: header },
          { value: "Qty", style: header },
          { value: "Due", style: header },
          { value: "Discount", style: header },
        ],
        ["A-1001", "Open", "High", 2, new Date(Date.UTC(2026, 11, 1)), 0.1],
        ["A-1002", "Closed", "Low", 10, new Date(Date.UTC(2026, 11, 15)), 0],
      ],
      validations: [
        {
          range: "B2:B100",
          list: ["Open", "Closed", "On hold"],
          input: { title: "Status", message: "Pick one of the three" },
        },
        { range: "C2:C100", list: "Lists!A2:A4" },
        {
          range: "D2:D100",
          number: { min: 1, max: 1000, integer: true },
          error: { title: "Quantity", message: "Enter a whole number from 1 to 1000" },
        },
        {
          range: "E2:E100",
          date: { min: new Date(Date.UTC(2026, 0, 1)), max: new Date(Date.UTC(2026, 11, 31)) },
          allowBlank: false,
          error: { message: "Due dates fall in 2026", style: "warning" },
        },
        {
          range: "F:F",
          number: { max: 0.5 },
          error: { message: "At most 50 %", style: "information" },
        },
      ],
    },
    { name: "Lists", rows: [["Priority"], ["High"], ["Medium"], ["Low"]] },
  ],
};

/** The workbook behind print.xlsx: landscape A4 fitted to one page wide, custom margins, header and footer, a repeated header row. */
/**
 * The workbook behind auto-width.xlsx: every column but one fits itself, with a bold header, a
 * merged title that must not count, money and date formats, booleans and a note past the ceiling.
 */
export const autoWidthWorkbook: Workbook = {
  sheets: [
    {
      name: "Inventory",
      autoWidth: true,
      columns: [
        undefined,
        undefined,
        { style: money },
        { style: { numberFormat: "yyyy-mm-dd" } },
        { width: 9 },
      ],
      merges: ["A1:F1"],
      rows: [
        [{ value: "Inventory, a title wider than its first column", style: { bold: true } }],
        [
          { value: "SKU", style: header },
          { value: "Product", style: header },
          { value: "Unit price", style: header },
          { value: "Restocked", style: header },
          { value: "In stock", style: header },
          { value: "Note", style: header },
        ],
        [
          "SKU-100001",
          "Height-adjustable desk",
          7490,
          new Date(Date.UTC(2026, 2, 14)),
          true,
          "Delivered and assembled on site",
        ],
        [
          "SKU-7",
          "Chair",
          1299.5,
          new Date(Date.UTC(2026, 9, 30)),
          false,
          "A note long enough to reach the sixty-character ceiling, which a fixed width would need to raise",
        ],
        ["SKU-100003", "Monitor arm, dual", 899, new Date(Date.UTC(2025, 11, 1)), true, ""],
        ["SKU-42", "Lamp", 249.9, new Date(Date.UTC(2026, 6, 7)), false, null],
        [{ value: "Total", style: header }, null, { formula: "SUM(C3:C6)", value: 9938.4 }],
      ],
    },
  ],
};

/**
 * The workbook behind images.xlsx: a picture per row in one column, the same badge repeated in
 * another, one picture scaled, a header row, a text column and a fixed-width column.
 */
export const imagesWorkbook: Workbook = {
  sheets: [
    {
      name: "Catalogue",
      autoWidth: true,
      columns: [undefined, { width: 10 }],
      rows: [
        [
          { value: "Product", style: header },
          { value: "Code", style: header },
          { value: "Picture", style: header },
          { value: "Badge", style: header },
        ],
        ["Desk", "D-1001", { image: stripesPng }, { image: badgePng }],
        ["Chair", "C-2002", { image: stripesPng, width: 60 }, { image: badgePng, height: 24 }],
        ["Lamp", "L-3003", { image: stripesPng, style: { fill: "#fff2cc" } }, { image: badgePng }],
        ["No picture", "N-0000", null, null],
      ],
    },
  ],
};

export const printWorkbook: Workbook = {
  sheets: [
    {
      name: "Report",
      columns: [{ width: 10 }, { width: 24 }, { width: 14, style: money }],
      rows: [
        [
          { value: "Month", style: header },
          { value: "Region", style: header },
          { value: "Revenue", style: header },
        ],
        ...Array.from({ length: 30 }, (_, i) => [
          `2026-${String((i % 12) + 1).padStart(2, "0")}`,
          ["North", "South", "East", "West"][i % 4],
          (i + 1) * 1250,
        ]),
      ],
      print: {
        orientation: "landscape",
        paperSize: "A4",
        fitTo: { width: 1 },
        margins: { left: 0.5, right: 0.5 },
        header: "&LQuarterly report&R&D",
        footer: "&L&A&RPage &P of &N",
        repeatRows: 1,
      },
    },
  ],
};
