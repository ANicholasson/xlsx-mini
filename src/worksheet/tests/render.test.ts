import { describe, expect, it } from "vite-plus/test";
import { XlsxMiniError } from "../../errors/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { type Cell, type Sheet, type WorksheetContext, renderWorksheet } from "../index.js";

/** Lets a test reach the runtime error branch for a value the types already reject. */
const unsafe = (value: unknown): Cell => value as Cell;

const styles = { dateOnly: () => 1, dateTime: () => 2 };

function render(sheet: Sheet, overrides: Partial<WorksheetContext> = {}): string {
  return renderWorksheet(sheet, {
    strings: createStringTable(),
    styles,
    coerce: undefined,
    sheetIndex: 0,
    ...overrides,
  });
}

const plainDate = {
  [Symbol.toStringTag]: "Temporal.PlainDate",
  year: 2026,
  month: 11,
  day: 4,
} as const;
const plainDateTime = {
  [Symbol.toStringTag]: "Temporal.PlainDateTime",
  year: 2000,
  month: 1,
  day: 1,
  hour: 12,
  minute: 0,
  second: 0,
  millisecond: 0,
  microsecond: 0,
  nanosecond: 0,
} as const;
const zoned = {
  ...plainDateTime,
  [Symbol.toStringTag]: "Temporal.ZonedDateTime",
  hour: 6,
} as const;
const instant = {
  [Symbol.toStringTag]: "Temporal.Instant",
  epochMilliseconds: Date.UTC(2000, 0, 1, 12),
} as const;

describe("renderWorksheet", () => {
  it("writes strings as shared-string references and numbers inline", () => {
    const strings = createStringTable();
    const xml = render(
      {
        name: "Sheet1",
        rows: [
          ["Name", "Age"],
          ["Kari", 42.5],
        ],
      },
      { strings },
    );
    expect(xml).toContain('<dimension ref="A1:B2"/>');
    expect(xml).toContain(
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>',
    );
    expect(xml).toContain('<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>42.5</v></c></row>');
  });

  it("writes booleans as boolean cells", () => {
    const xml = render({ name: "S", rows: [[true, false]] });
    expect(xml).toContain('<c r="A1" t="b"><v>1</v></c><c r="B1" t="b"><v>0</v></c>');
  });

  it("writes a Date as a UTC serial with the date-time style", () => {
    const xml = render({ name: "S", rows: [[new Date(Date.UTC(2000, 0, 1, 12))]] });
    expect(xml).toContain('<c r="A1" s="2"><v>36526.5</v></c>');
  });

  it("writes Temporal-like values by their tag, date-only ones with the date style", () => {
    const xml = render({ name: "S", rows: [[plainDate, plainDateTime, zoned, instant]] });
    expect(xml).toContain('<c r="A1" s="1"><v>46330</v></c>');
    expect(xml).toContain('<c r="B1" s="2"><v>36526.5</v></c>');
    expect(xml).toContain('<c r="C1" s="2"><v>36526.25</v></c>');
    expect(xml).toContain('<c r="D1" s="2"><v>36526.5</v></c>');
  });

  it("unwraps a plain { value } object once", () => {
    const xml = render({ name: "S", rows: [[{ value: 7 }, { value: null }, { value: "x" }]] });
    expect(xml).toContain('<row r="1"><c r="A1"><v>7</v></c><c r="C1" t="s"><v>0</v></c></row>');
  });

  it("skips null and undefined cells but keeps their position", () => {
    const xml = render({ name: "S", rows: [[null, undefined, 7]] });
    expect(xml).toContain('<row r="1"><c r="C1"><v>7</v></c></row>');
    expect(xml).toContain('<dimension ref="A1:C1"/>');
  });

  it("renders an empty sheet with an empty sheetData element", () => {
    const xml = render({ name: "S", rows: [] });
    expect(xml).toContain("<sheetData/>");
    expect(xml).not.toContain("<dimension");
  });

  it("accepts any iterable of rows", () => {
    function* rows() {
      yield [1];
      yield [2];
    }
    const xml = render({ name: "S", rows: rows() });
    expect(xml).toContain('<row r="2"><c r="A2"><v>2</v></c></row>');
  });

  describe("coerce", () => {
    it("receives every non-primitive value once and its result is handled normally", () => {
      const seen: string[] = [];
      const coerce: WorksheetContext["coerce"] = (value, path) => {
        seen.push(path);
        if (typeof value === "bigint") return Number(value);
        if (value instanceof Date) return plainDate;
        return value;
      };
      const xml = render(
        { name: "S", rows: [[10n, new Date(0), "text", 1, plainDateTime]] },
        { coerce, sheetIndex: 2 },
      );
      expect(seen).toEqual([
        "sheets[2].rows[0][0]",
        "sheets[2].rows[0][1]",
        "sheets[2].rows[0][4]",
      ]);
      expect(xml).toContain('<c r="A1"><v>10</v></c>');
      expect(xml).toContain('<c r="B1" s="1"><v>46330</v></c>');
      expect(xml).toContain('<c r="E1" s="2"><v>36526.5</v></c>');
    });

    it("is not applied to its own result", () => {
      let calls = 0;
      const coerce: WorksheetContext["coerce"] = () => {
        calls++;
        return new Date(Date.UTC(2000, 0, 1));
      };
      render({ name: "S", rows: [[new Date(0)]] }, { coerce });
      expect(calls).toBe(1);
    });
  });

  describe("errors", () => {
    const failure = (sheet: Sheet, overrides: Partial<WorksheetContext> = {}) => {
      try {
        render(sheet, overrides);
      } catch (error) {
        return error as XlsxMiniError;
      }
      throw new Error("expected a throw");
    };

    it("rejects NaN and infinite numbers with the cell path", () => {
      expect(failure({ name: "S", rows: [[1], [2, Number.NaN]] })).toMatchObject({
        code: "INVALID_NUMBER",
        path: "sheets[0].rows[1][1]",
      });
      expect(
        failure({ name: "S", rows: [[Number.POSITIVE_INFINITY]] }, { sheetIndex: 3 }),
      ).toMatchObject({
        code: "INVALID_NUMBER",
        path: "sheets[3].rows[0][0]",
      });
    });

    it("rejects dates outside Excel's range and invalid dates", () => {
      expect(failure({ name: "S", rows: [[new Date(Date.UTC(1899, 0, 1))]] })).toMatchObject({
        code: "DATE_OUT_OF_RANGE",
        path: "sheets[0].rows[0][0]",
      });
      expect(failure({ name: "S", rows: [[new Date(Number.NaN)]] })).toMatchObject({
        code: "INVALID_DATE",
      });
      expect(failure({ name: "S", rows: [[{ ...plainDate, year: 1850 }]] })).toMatchObject({
        code: "DATE_OUT_OF_RANGE",
      });
    });

    it("rejects values it cannot write, including unknown Temporal types and nested wrappers", () => {
      expect(
        failure({ name: "S", rows: [[unsafe({ [Symbol.toStringTag]: "Temporal.Duration" })]] }),
      ).toMatchObject({
        code: "UNSUPPORTED_VALUE",
      });
      expect(failure({ name: "S", rows: [[unsafe(new Map())]] })).toMatchObject({
        code: "UNSUPPORTED_VALUE",
      });
      expect(failure({ name: "S", rows: [[10n]] })).toMatchObject({
        code: "UNSUPPORTED_VALUE",
        path: "sheets[0].rows[0][0]",
      });
      expect(failure({ name: "S", rows: [[unsafe({ value: { value: 1 } })]] })).toMatchObject({
        code: "UNSUPPORTED_VALUE",
      });
      expect(failure({ name: "S", rows: [[unsafe(Symbol("s"))]] })).toMatchObject({
        code: "UNSUPPORTED_VALUE",
      });
    });

    it("enforces Excel's limits on text length, columns and rows", () => {
      expect(failure({ name: "S", rows: [["x".repeat(32768)]] })).toMatchObject({
        code: "CELL_TEXT_TOO_LONG",
      });
      expect(render({ name: "S", rows: [["x".repeat(32767)]] })).toContain("<v>0</v>");
      const wide = Array.from({ length: 16385 }, () => 1);
      expect(failure({ name: "S", rows: [wide] })).toMatchObject({
        code: "TOO_MANY_COLUMNS",
        path: "sheets[0].rows[0]",
      });
      function* tooMany() {
        for (let i = 0; i < 1_048_577; i++) yield [];
      }
      expect(failure({ name: "S", rows: tooMany() })).toMatchObject({
        code: "TOO_MANY_ROWS",
        path: "sheets[0].rows[1048576]",
      });
    });
  });
});
