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
