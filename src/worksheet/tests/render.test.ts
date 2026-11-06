import { describe, expect, it } from "vite-plus/test";
import { XlsxMiniError } from "../../errors/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { type Cell, type Sheet, type WorksheetContext, renderWorksheet } from "../index.js";

/** Lets a test reach the runtime error branch for a value the types already reject. */
const unsafe = (value: unknown): Cell => value as Cell;

import { createStyleTable } from "../../styles/index.js";

/** A real style table: the worksheet's contract with it is exercised, not stubbed. */
const styles = createStyleTable();

function render(sheet: Sheet, overrides: Partial<WorksheetContext> = {}): string {
  const rendered = renderWorksheet(sheet, {
    strings: createStringTable(),
    styles,
    coerce: undefined,
    sheetIndex: 0,
    ...overrides,
  });
  return new TextDecoder().decode(rendered.data);
}

const failure = (sheet: Sheet, overrides: Partial<WorksheetContext> = {}): XlsxMiniError => {
  try {
    render(sheet, overrides);
  } catch (error) {
    return error as XlsxMiniError;
  }
  throw new Error("expected a throw");
};

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
    const table = createStyleTable();
    const xml = render(
      { name: "S", rows: [[new Date(Date.UTC(2000, 0, 1, 12))]] },
      { styles: table },
    );
    expect(xml).toContain('<c r="A1" s="1"><v>36526.5</v></c>');
    expect(table.render()).toContain('numFmtId="22"');
  });

  it("writes Temporal-like values by their tag, date-only ones with the date style", () => {
    const table = createStyleTable();
    const xml = render(
      { name: "S", rows: [[plainDate, plainDateTime, zoned, instant]] },
      { styles: table },
    );
    expect(xml).toContain('<c r="A1" s="1"><v>46330</v></c>');
    expect(xml).toContain('<c r="B1" s="2"><v>36526.5</v></c>');
    expect(xml).toContain('<c r="C1" s="2"><v>36526.25</v></c>');
    expect(xml).toContain('<c r="D1" s="2"><v>36526.5</v></c>');
    expect(table.render()).toContain('<cellXfs count="3">');
  });

  describe("styles", () => {
    it("writes the cell format index for styled cells and nothing for unstyled ones", () => {
      const table = createStyleTable();
      const xml = render(
        {
          name: "S",
          rows: [
            [
              { value: "Total", style: { bold: true } },
              { value: 1.5, style: { numberFormat: "0.00" } },
              2,
            ],
          ],
        },
        { styles: table },
      );
      expect(xml).toContain(
        '<c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="2"><v>1.5</v></c><c r="C1"><v>2</v></c>',
      );
    });

    it("applies a column style to every cell in the column that has none of its own, cell properties winning", () => {
      const table = createStyleTable();
      const sheet: Sheet = {
        name: "S",
        columns: [undefined, { style: { numberFormat: "0.0%" } }],
        rows: [
          ["Share", { value: "Share", style: { bold: true } }],
          [0.25, 0.25],
          [
            { value: 0.5, style: { numberFormat: "0%" } },
            { value: 0.5, style: { numberFormat: "0%" } },
          ],
          [undefined, undefined, 7],
        ],
      };
      const xml = render(sheet, { styles: table });
      expect(xml).toContain('<c r="A1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>0</v></c>');
      expect(xml).toContain('<c r="A2"><v>0.25</v></c><c r="B2" s="2"><v>0.25</v></c>');
      expect(xml).toContain('<c r="A3" s="3"><v>0.5</v></c><c r="B3" s="3"><v>0.5</v></c>');
      expect(xml).toContain('<c r="C4"><v>7</v></c>');
      expect(table.render()).toContain('<cellXfs count="4">');
    });

    it("lets an explicit number format override the date default, from the cell or the column", () => {
      const table = createStyleTable();
      const date = new Date(Date.UTC(2026, 10, 4));
      const xml = render(
        {
          name: "S",
          columns: [{ style: { numberFormat: "yyyy-mm-dd" } }],
          rows: [[date, date, { value: date, style: { numberFormat: 15 } }]],
        },
        { styles: table },
      );
      expect(xml).toContain(
        '<c r="A1" s="1"><v>46330</v></c><c r="B1" s="2"><v>46330</v></c><c r="C1" s="3"><v>46330</v></c>',
      );
      const stylesXml = table.render();
      expect(stylesXml).toContain('formatCode="yyyy-mm-dd"');
      expect(stylesXml).toContain('numFmtId="22"');
      expect(stylesXml).toContain('numFmtId="15"');
    });

    it("writes no style attribute when a style resolves to the default format", () => {
      const table = createStyleTable();
      const xml = render(
        {
          name: "S",
          columns: [{ style: {} }],
          rows: [[{ value: 1, style: { bold: false } }, "x"]],
        },
        { styles: table },
      );
      expect(xml).toContain('<row r="1"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>0</v></c></row>');
      expect(table.render()).toContain('<cellXfs count="1">');
    });

    it("applies a bold-only column style and validates it without registering a format", () => {
      const table = createStyleTable();
      const xml = render(
        { name: "S", columns: [{ style: { bold: true } }], rows: [["x", "y"]] },
        { styles: table },
      );
      expect(xml).toContain('<c r="A1" s="1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>');
      expect(table.render()).not.toContain("<numFmts");
    });

    it("lets an unexpected error from the styles port pass through untouched", () => {
      const boom = new Error("disk on fire");
      const broken = {
        indexFor: () => {
          throw boom;
        },
        validate: () => {
          throw boom;
        },
      };
      expect(() =>
        render({ name: "S", rows: [[{ value: 1, style: { bold: true } }]] }, { styles: broken }),
      ).toThrow(boom);
      expect(() =>
        render({ name: "S", columns: [{ style: { bold: true } }], rows: [] }, { styles: broken }),
      ).toThrow(boom);
    });

    it("reports an invalid number format at the cell or the column that carries it", () => {
      expect(
        failure({ name: "S", rows: [[1], [{ value: 2, style: { numberFormat: "" } }]] }),
      ).toMatchObject({
        code: "INVALID_NUMBER_FORMAT",
        path: "sheets[0].rows[1][0].style.numberFormat",
      });
      expect(
        failure(
          { name: "S", columns: [undefined, { style: { numberFormat: 99 } }], rows: [[1, 2]] },
          { sheetIndex: 4 },
        ),
      ).toMatchObject({
        code: "INVALID_NUMBER_FORMAT",
        path: "sheets[4].columns[1].style.numberFormat",
      });
    });
  });

  it("unwraps a plain { value } object once", () => {
    const xml = render({ name: "S", rows: [[{ value: 7 }, { value: null }, { value: "x" }]] });
    expect(xml).toContain('<row r="1"><c r="A1"><v>7</v></c><c r="C1" t="s"><v>0</v></c></row>');
  });

  it("treats a null-prototype object with a value property as a cell object too", () => {
    const bare = Object.assign(Object.create(null) as Record<string, unknown>, { value: 3 });
    expect(render({ name: "S", rows: [[unsafe(bare)]] })).toContain('<c r="A1"><v>3</v></c>');
    const nothing = Object.create(null) as Record<string, unknown>;
    expect(failure({ name: "S", rows: [[unsafe(nothing)]] })).toMatchObject({
      code: "UNSUPPORTED_VALUE",
    });
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

  describe("layout", () => {
    it("writes column widths in character units only for columns that set one", () => {
      const xml = render({
        name: "S",
        columns: [{ width: 12 }, undefined, { width: 32.5, style: { bold: true } }],
        rows: [[1, 2, 3]],
      });
      expect(xml).toContain(
        '<cols><col min="1" max="1" width="12" customWidth="1"/><col min="3" max="3" width="32.5" customWidth="1"/></cols>',
      );
      expect(xml.indexOf("<cols>")).toBeLessThan(xml.indexOf("<sheetData>"));
    });

    it("writes no cols element when no column sets a width", () => {
      expect(
        render({ name: "S", columns: [{ style: { bold: true } }], rows: [[1]] }),
      ).not.toContain("<cols>");
    });

    it("freezes rows, columns, or both, with the pane Excel expects", () => {
      expect(render({ name: "S", freeze: { rows: 1 }, rows: [[1]] })).toContain(
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
      );
      expect(render({ name: "S", freeze: { columns: 2 }, rows: [[1]] })).toContain(
        '<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>',
      );
      expect(render({ name: "S", freeze: { rows: 3, columns: 1 }, rows: [[1]] })).toContain(
        '<pane xSplit="1" ySplit="3" topLeftCell="B4" activePane="bottomRight" state="frozen"/>',
      );
      expect(render({ name: "S", freeze: { rows: 0, columns: 0 }, rows: [[1]] })).not.toContain(
        "<sheetViews>",
      );
    });

    it("puts the autofilter over the used range, after the data, and skips it on an empty sheet", () => {
      const xml = render({
        name: "S",
        autoFilter: true,
        rows: [
          ["a", "b", "c"],
          [1, 2, 3],
          [4, 5, 6],
        ],
      });
      expect(xml).toContain('</sheetData><autoFilter ref="A1:C3"/></worksheet>');
      expect(render({ name: "S", autoFilter: true, rows: [] })).not.toContain("<autoFilter");
    });

    it("keeps the schema order: dimension, sheetViews, cols, sheetData, autoFilter", () => {
      const xml = render({
        name: "S",
        columns: [{ width: 9 }],
        freeze: { rows: 1 },
        autoFilter: true,
        rows: [["h"], [1]],
      });
      const order = ["<dimension", "<sheetViews>", "<cols>", "<sheetData>", "<autoFilter"].map(
        (tag) => xml.indexOf(tag),
      );
      expect(order).toEqual([...order].sort((a, b) => a - b));
      expect(order.every((i) => i >= 0)).toBe(true);
    });

    it("rejects widths and freeze counts Excel would refuse, at their path", () => {
      for (const width of [0, -1, 256, Number.NaN, Number.POSITIVE_INFINITY]) {
        expect(
          failure({ name: "S", columns: [undefined, { width }], rows: [] }, { sheetIndex: 1 }),
        ).toMatchObject({
          code: "INVALID_COLUMN_WIDTH",
          path: "sheets[1].columns[1].width",
        });
      }
      expect(failure({ name: "S", freeze: { rows: -1 }, rows: [] })).toMatchObject({
        code: "INVALID_FREEZE",
        path: "sheets[0].freeze.rows",
      });
      expect(failure({ name: "S", freeze: { rows: 1.5 }, rows: [] })).toMatchObject({
        code: "INVALID_FREEZE",
        path: "sheets[0].freeze.rows",
      });
      expect(failure({ name: "S", freeze: { columns: 16384 }, rows: [] })).toMatchObject({
        code: "INVALID_FREEZE",
        path: "sheets[0].freeze.columns",
      });
      expect(failure({ name: "S", freeze: { rows: 1048576 }, rows: [] })).toMatchObject({
        code: "INVALID_FREEZE",
        path: "sheets[0].freeze.rows",
      });
    });
  });

  describe("errors", () => {
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

  it("writes no cell when coerce returns null or undefined", () => {
    const sheet: Sheet = { name: "S", rows: [[new Date(0), "kept"]] };
    expect(render(sheet, { coerce: () => null })).toContain('<row r="1"><c r="B1" t="s">');
    expect(render(sheet, { coerce: () => undefined })).not.toContain('r="A1"');
  });

  describe("appearance", () => {
    it("writes a styled empty cell for a cell object with a style and no value, and nothing for a bare null in a styled column", () => {
      const xml = render({
        name: "S",
        columns: [{ style: { fill: "#eeeeee" } }],
        rows: [
          [
            null,
            { value: null, style: { border: "thin" } },
            { value: undefined, style: {} },
            { value: null },
          ],
        ],
      });
      expect(xml).toMatch(/<row r="1"><c r="B1" s="\d+"\/><\/row>/);
      expect(xml).toContain('<dimension ref="A1:B1"/>');
    });

    it("reports an invalid colour, border, alignment or indent at the property that carries it", () => {
      expect(
        failure({ name: "S", rows: [[{ value: 1, style: { fontColor: "red" } }]] }),
      ).toMatchObject({ code: "INVALID_STYLE", path: "sheets[0].rows[0][0].style.fontColor" });
      expect(
        failure({
          name: "S",
          columns: [{ style: { border: { top: unsafe("dashed") as never } } }],
          rows: [],
        }),
      ).toMatchObject({ code: "INVALID_STYLE", path: "sheets[0].columns[0].style.border.top" });
      expect(
        failure({
          name: "S",
          columns: [{ style: { align: unsafe("justify") as never } }],
          rows: [],
        }),
      ).toMatchObject({ code: "INVALID_STYLE", path: "sheets[0].columns[0].style.align" });
      expect(failure({ name: "S", rows: [[{ value: 1, style: { indent: 1.5 } }]] })).toMatchObject({
        code: "INVALID_STYLE",
        path: "sheets[0].rows[0][0].style.indent",
      });
    });
  });

  describe("merges", () => {
    it("writes the merged ranges after the sheet data and the autofilter", () => {
      const xml = render({
        name: "S",
        autoFilter: true,
        merges: ["A1:C1", "A3:A4"],
        rows: [["Title"], ["a", "b", "c"]],
      });
      expect(xml).toContain(
        '</sheetData><autoFilter ref="A1:C2"/><mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="A3:A4"/></mergeCells></worksheet>',
      );
    });

    it("writes no element for an empty list", () => {
      expect(render({ name: "S", merges: [], rows: [["a"]] })).not.toContain("mergeCells");
    });

    it("rejects ranges that are not A1 form, reversed, single cells, outside the sheet or overlapping", () => {
      const bad = (merges: string[]) => failure({ name: "S", merges, rows: [] });
      expect(bad(["A1"])).toMatchObject({ code: "INVALID_MERGE", path: "sheets[0].merges[0]" });
      expect(bad(["a1:b1"])).toMatchObject({ code: "INVALID_MERGE" });
      expect(bad(["A1:B1", "C1:A2"])).toMatchObject({
        code: "INVALID_MERGE",
        path: "sheets[0].merges[1]",
      });
      expect(bad(["B2:B2"])).toMatchObject({ code: "INVALID_MERGE" });
      expect(bad(["XFE1:XFF1"])).toMatchObject({ code: "INVALID_MERGE" });
      expect(bad(["A1048576:A1048577"])).toMatchObject({ code: "INVALID_MERGE" });
      expect(bad(["A1:C3", "C3:D4"])).toMatchObject({
        code: "INVALID_MERGE",
        path: "sheets[0].merges[1]",
      });
      expect(bad([unsafe(7) as never])).toMatchObject({ code: "INVALID_MERGE" });
      expect(bad([unsafe(10n) as never])).toMatchObject({
        code: "INVALID_MERGE",
        path: "sheets[0].merges[0]",
        message: 'A merge is a range in A1 form such as "A1:D1", not 10n at sheets[0].merges[0]',
      });
      expect(render({ name: "S", merges: ["A1:C3", "D3:D4", "XFD1:XFD2"], rows: [] })).toContain(
        '<mergeCell ref="XFD1:XFD2"/>',
      );
    });
  });

  describe("row heights", () => {
    it("writes a height on the row that has one, and an empty row for a height without cells", () => {
      const xml = render({
        name: "S",
        rowHeights: { 1: 28, 2: 15.5, 5: 40 },
        rows: [["title"], [], ["a"]],
      });
      expect(xml).toContain(
        '<row r="1" ht="28" customHeight="1"><c r="A1" t="s"><v>0</v></c></row>',
      );
      expect(xml).toContain('<row r="2" ht="15.5" customHeight="1"/>');
      expect(xml).toContain('<row r="3"><c r="A3" t="s"><v>1</v></c></row>');
      expect(xml).toContain('</row><row r="5" ht="40" customHeight="1"/></sheetData>');
      expect(xml).toContain('<dimension ref="A1:A3"/>');
    });

    it("rejects a style that is not an object, at the cell or the column", () => {
      expect(failure({ name: "S", rows: [[unsafe({ value: 1, style: null })]] })).toMatchObject({
        code: "INVALID_STYLE",
        path: "sheets[0].rows[0][0].style",
      });
      expect(
        failure({ name: "S", rows: [[unsafe({ value: null, style: "bold" })]] }),
      ).toMatchObject({ code: "INVALID_STYLE", path: "sheets[0].rows[0][0].style" });
      expect(
        failure({ name: "S", columns: [{ style: null as never }], rows: [["a"]] }),
      ).toMatchObject({ code: "INVALID_STYLE", path: "sheets[0].columns[0].style" });
    });

    it("keeps height-only rows inside sheetData when the sheet has no cells", () => {
      const xml = render({ name: "S", rowHeights: { 1: 28 }, rows: [] });
      expect(xml).toContain('<sheetData><row r="1" ht="28" customHeight="1"/></sheetData>');
      expect(xml).not.toContain("<sheetData/>");
      expect(xml).not.toContain("<dimension");
      const sparse = render({ name: "S", rowHeights: { 2: 28 }, rows: [[null], []] });
      expect(sparse).toContain('<sheetData><row r="2" ht="28" customHeight="1"/></sheetData>');
    });

    it("writes nothing extra when every height belongs to a row that was given", () => {
      const xml = render({ name: "S", rowHeights: { 1: 20 }, rows: [["a"], ["b"]] });
      expect(xml).toContain('<row r="1" ht="20" customHeight="1">');
      expect(xml.match(/<row /g)).toHaveLength(2);
    });

    it("writes trailing heights in row order whatever the key order", () => {
      const xml = render({ name: "S", rowHeights: { 9: 10, 4: 20 }, rows: [["a"]] });
      expect(xml).toContain(
        '<row r="4" ht="20" customHeight="1"/><row r="9" ht="10" customHeight="1"/>',
      );
    });

    it("rejects row numbers below 1 or beyond the sheet and heights outside 0 to 409", () => {
      const bad = (rowHeights: Record<number, number>) =>
        failure({ name: "S", rowHeights, rows: [] });
      expect(bad({ 0: 10 })).toMatchObject({
        code: "INVALID_ROW_HEIGHT",
        path: "sheets[0].rowHeights[0]",
      });
      expect(bad({ 1048577: 10 })).toMatchObject({ code: "INVALID_ROW_HEIGHT" });
      expect(bad({ 2: 0 })).toMatchObject({
        code: "INVALID_ROW_HEIGHT",
        path: "sheets[0].rowHeights[2]",
      });
      expect(bad({ 2: 410 })).toMatchObject({ code: "INVALID_ROW_HEIGHT" });
      expect(bad({ 2: Number.NaN })).toMatchObject({ code: "INVALID_ROW_HEIGHT" });
      expect(bad({ [unsafe("x") as never]: 10 })).toMatchObject({ code: "INVALID_ROW_HEIGHT" });
    });
  });
});
