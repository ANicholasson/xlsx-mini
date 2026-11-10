import { describe, expect, it } from "vite-plus/test";
import type { XlsxMiniError } from "../../errors/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import { type Cell, type Sheet, renderWorksheet } from "../index.js";

const plainDate = {
  [Symbol.toStringTag]: "Temporal.PlainDate",
  year: 2026,
  month: 11,
  day: 4,
} as const;

function render(sheet: Sheet): string {
  const rendered = renderWorksheet(sheet, {
    strings: createStringTable(),
    styles: createStyleTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["S"],
  });
  return new TextDecoder().decode(rendered.data);
}

/** The widths of the `<cols>` element by column position, or an empty record. */
function cols(sheet: Sheet): Record<number, number> {
  const widths: Record<number, number> = {};
  for (const match of render(sheet).matchAll(/<col min="(\d+)" max="\d+" width="([\d.]+)"/g)) {
    widths[Number(match[1])] = Number(match[2]);
  }
  return widths;
}

/** The auto width of column A holding these cells, one per row, or undefined when none is written. */
function width(cells: readonly Cell[], sheet: Partial<Sheet> = {}): number | undefined {
  return cols({
    name: "S",
    columns: [{ width: "auto" }],
    rows: cells.map((cell) => [cell]),
    ...sheet,
  })[1];
}

describe("auto column width", () => {
  describe("text", () => {
    it("measures Calibri 11 character by character, so the same count can differ in width", () => {
      expect(width(["Wim"])).toBe(5.71);
      expect(width(["illi"])).toBe(3.43);
    });

    it("takes the widest cell, however the rows are ordered", () => {
      expect(width(["Kari Nordmann", "Oslo", "Ola"])).toBe(14.86);
      expect(width(["Ola", "Oslo", "Kari Nordmann"])).toBe(14.86);
    });

    it("makes bold a tenth wider and scales with the font size, from the cell or the column", () => {
      expect(width(["Id"])).toBe(3.43);
      expect(width([{ value: "Id", style: { bold: true } }])).toBe(3.6);
      expect(width([{ value: "Id", style: { fontSize: 22 } }])).toBe(5.14);
      expect(width(["Id"], { columns: [{ width: "auto", style: { bold: true } }] })).toBe(3.6);
      expect(
        width([{ value: "Id", style: { bold: false } }], {
          columns: [{ width: "auto", style: { bold: true } }],
        }),
      ).toBe(3.43);
    });

    it("measures the widest line of a text with newlines", () => {
      expect(width(["ab\ncdefgh"])).toBe(7.71);
      expect(width(["cdefgh"])).toBe(7.71);
      expect(width(["cdefgh\nab\nc"])).toBe(7.71);
    });

    it("counts wide characters double, other characters as eight pixels, marks and controls as nothing", () => {
      expect(width(["日本"])).toBe(5.71);
      expect(width(["🎉"])).toBe(3.71);
      // One character from each wide range: Hangul Jamo, CJK, Hangul syllables, compatibility
      // ideographs, vertical forms, fullwidth forms and fullwidth signs.
      expect(width(["\u1100\u65e5\uac00\uf900\ufe30\uff01\uffe0"])).toBe(15.71);
      expect(width(["é"])).toBe(3);
      expect(width(["é"])).toBe(3);
      expect(width(["\t"])).toBeUndefined();
    });

    it("measures booleans as TRUE and FALSE, and a link's text", () => {
      expect(width([true])).toBe(6.14);
      expect(width([false])).toBe(6.86);
      expect(width([{ value: "Docs", href: "https://example.com" }])).toBe(5.86);
    });

    it("stays between 3 and 60 characters, and stops measuring at the ceiling", () => {
      expect(width(["."])).toBe(3);
      expect(width(["x".repeat(500)])).toBe(60);
      expect(width(["x".repeat(500), "y".repeat(5000)])).toBe(60);
    });
  });

  describe("numbers", () => {
    it("measures the General format as the number's text, at most eleven characters", () => {
      expect(width([12345])).toBe(6.71);
      expect(width([-12345])).toBe(7.71);
      expect(width([0])).toBe(3);
      expect(width([0.5])).toBe(4.71);
      expect(width([1234567.891234])).toBe(12.71);
      expect(width([1e21])).toBe(6.71);
    });

    it("reads a format's placeholders, grouping, decimals and literal text", () => {
      const money = { numberFormat: '#,##0.00 "kr"' };
      expect(width([{ value: -1234567.891, style: money }])).toBe(16.86);
      expect(width([{ value: 5, style: money }])).toBe(7.86);
      expect(width([5], { columns: [{ width: "auto", style: money }] })).toBe(7.86);
      expect(
        width([
          { value: 5, style: money },
          { value: -1234567.891, style: money },
        ]),
      ).toBe(16.86);
      // A cell style without a format falls back to the column's, or to General.
      expect(
        width([{ value: 5, style: { bold: true } }], {
          columns: [{ width: "auto", style: money }],
        }),
      ).toBe(8.47);
      expect(width([{ value: 5, style: { bold: true } }])).toBe(3);
    });

    it("scales by percent and by a trailing comma, and shows an exponent as a short mantissa", () => {
      expect(width([{ value: 0.153, style: { numberFormat: "0%" } }])).toBe(5.29);
      expect(width([{ value: 123456, style: { numberFormat: "0.00E+00" } }])).toBe(9.71);
      expect(width([{ value: 123456, style: { numberFormat: "0.00e-00" } }])).toBe(9.71);
      expect(width([{ value: 1234567, style: { numberFormat: "#,##0," } }])).toBe(6.71);
    });

    it("reads padding, fill, colour, escapes, a text placeholder and only the first section", () => {
      expect(width([{ value: 1234, style: { numberFormat: "_(#,##0_);[Red](#,##0)" } }])).toBe(
        7.57,
      );
      expect(width([{ value: 7, style: { numberFormat: "* 0" } }])).toBe(3);
      expect(width([{ value: 7, style: { numberFormat: "[Red" } }])).toBe(3);
      expect(width([{ value: 0, style: { numberFormat: "0.00" } }])).toBe(5.71);
      expect(width([{ value: 7, style: { numberFormat: "\\k0" } }])).toBe(3.71);
      expect(width([{ value: 7, style: { numberFormat: "0\\" } }])).toBe(3);
      expect(width([{ value: 5, style: { numberFormat: '"abc' } }])).toBe(5.71);
      expect(width([{ value: 5, style: { numberFormat: "0.00," } }])).toBe(6.29);
      expect(width([{ value: 5, style: { numberFormat: "0.0.0" } }])).toBe(6.29);
      expect(width([{ value: 5, style: { numberFormat: "@" } }])).toBe(3);
    });

    it("knows the built-in formats that show something other than General", () => {
      expect(width([{ value: 1234.5, style: { numberFormat: 4 } }])).toBe(9.71);
      expect(width([{ value: 0.5, style: { numberFormat: 9 } }])).toBe(5.29);
      expect(width([{ value: 1234.5, style: { numberFormat: 5 } }])).toBe(7.71);
      expect(width([{ value: 1234.5, style: { numberFormat: 0 } }])).toBe(7.71);
    });
  });

  describe("dates", () => {
    it("measures an unstyled date by its default format: m/d/yyyy for a plain date, with the time for a Date", () => {
      expect(width([plainDate])).toBe(11.43);
      expect(width([new Date(Date.UTC(2026, 10, 4, 13, 45))])).toBe(16.43);
    });

    it("reads a date format's tokens: month and day names, years, two digits for the rest", () => {
      expect(width([{ value: plainDate, style: { numberFormat: "yyyy-mm-dd" } }])).toBe(11.14);
      expect(width([{ value: plainDate, style: { numberFormat: "dddd d mmmm yyyy" } }])).toBe(27);
      expect(width([{ value: plainDate, style: { numberFormat: "h:mm AM/PM" } }])).toBe(8.71);
      expect(width([{ value: plainDate, style: { numberFormat: "h A/P" } }])).toBe(5.14);
      expect(width([{ value: plainDate, style: { numberFormat: "[h]:mm:ss" } }])).toBe(8.86);
      expect(width([{ value: plainDate, style: { numberFormat: "mmmmm" } }])).toBe(3);
      expect(width([{ value: plainDate, style: { numberFormat: "ddd mmm" } }])).toBe(8.14);
      expect(width([{ value: plainDate, style: { numberFormat: "e" } }])).toBe(3.71);
      expect(width([{ value: plainDate, style: { numberFormat: "bbbb" } }])).toBe(5.71);
      expect(width([{ value: plainDate, style: { numberFormat: "ggg" } }])).toBe(3.71);
    });

    it("follows the format rather than the value's kind: a number under a date format, a date under a numeric one", () => {
      expect(
        width([5], { columns: [{ width: "auto", style: { numberFormat: "yyyy-mm-dd" } }] }),
      ).toBe(11.14);
      expect(width([{ value: plainDate, style: { numberFormat: "0.00" } }])).toBe(9.71);
    });
  });

  describe("formulas and merges", () => {
    it("measures a formula's cached value as its own kind, and nothing without one", () => {
      expect(width([{ formula: "A1", value: "abc" }])).toBe(4.71);
      expect(width([{ formula: "A1", value: 12345 }])).toBe(6.71);
      expect(width([{ formula: "A1", value: true }])).toBe(6.14);
      expect(width([{ formula: "A1", value: false }])).toBe(6.86);
      expect(width([{ formula: "A1" }])).toBeUndefined();
    });

    it("skips the cell that starts a merge across columns, and counts one down a column", () => {
      expect(width(["A long merged title", "ab"], { merges: ["A1:C1"] })).toBe(3.86);
      expect(width(["A long merged title", "ab"], { merges: ["A1:A2"] })).toBe(18.29);
      expect(width(["A long merged title", "ab"])).toBe(18.29);
    });
  });

  describe("which columns", () => {
    it('sizes a column with width "auto", and writes no width for one whose cells are all empty', () => {
      expect(
        cols({
          name: "S",
          columns: [{ width: "auto" }, { width: "auto" }, { width: 9 }],
          rows: [
            ["ab", null, 1],
            ["", { value: null, style: { bold: true } }, 2],
          ],
        }),
      ).toEqual({ 1: 3.86, 3: 9 });
    });

    it("sizes every column without a numeric width under autoWidth, past the column list too", () => {
      expect(
        cols({
          name: "S",
          autoWidth: true,
          columns: [{ width: 9 }, undefined, { style: { bold: true } }],
          rows: [["ab", "ab", "ab", "ab", "cdefgh"]],
        }),
      ).toEqual({ 1: 9, 2: 3.86, 3: 4.07, 4: 3.86, 5: 7.71 });
      expect(cols({ name: "S", columns: [{ width: "auto" }], rows: [["ab", "cdefgh"]] })).toEqual({
        1: 3.86,
      });
    });

    it("keeps the schema order: cols between the views and the rows", () => {
      const xml = render({
        name: "S",
        autoWidth: true,
        freeze: { rows: 1 },
        rows: [["ab"]],
      });
      expect(xml).toContain(
        '</sheetViews><cols><col min="1" max="1" width="3.86" customWidth="1"/></cols><sheetData>',
      );
    });

    it("rejects any string but auto as a width, at the column's path", () => {
      let caught: XlsxMiniError | undefined;
      try {
        render({ name: "S", columns: [{ width: "wide" as never }], rows: [] });
      } catch (error) {
        caught = error as XlsxMiniError;
      }
      expect(caught).toMatchObject({
        code: "INVALID_COLUMN_WIDTH",
        path: "sheets[0].columns[0].width",
      });
      expect(caught?.message).toContain('or "auto"');
    });
  });
});
