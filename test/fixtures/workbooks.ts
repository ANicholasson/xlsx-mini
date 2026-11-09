import { type WriteOptions, type Workbook, sheetFromObjects } from "../../src/index.js";

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
