import type { WriteOptions, Workbook } from "../../src/index.js";

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
