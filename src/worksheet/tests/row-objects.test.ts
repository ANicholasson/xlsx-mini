import { describe, expect, it } from "vite-plus/test";
import { createMediaTable } from "../../drawing/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { type CellStyle, createStyleTable } from "../../styles/index.js";
import { type Row, type Sheet, type WorksheetContext, renderWorksheet } from "../index.js";

function context(overrides: Partial<WorksheetContext> = {}): WorksheetContext {
  return {
    strings: createStringTable(),
    media: createMediaTable(),
    styles: createStyleTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["S"],
    inlineStrings: true,
    ...overrides,
  };
}

const xml = (sheet: Sheet, overrides?: Partial<WorksheetContext>) =>
  new TextDecoder().decode(renderWorksheet(sheet, context(overrides)).data);

/** The `<sheetData>` of a sheet with these rows. */
const data = (rows: Row[], sheet: Partial<Sheet> = {}, overrides?: Partial<WorksheetContext>) =>
  /<sheetData>(.*)<\/sheetData>/.exec(xml({ name: "S", rows, ...sheet }, overrides))![1]!;

const bold: CellStyle = { bold: true };

describe("a row given as an object", () => {
  it("writes what the same cells write as an array when it carries nothing else", () => {
    const cells = ["a", 1, null, true, { value: null }, 2];
    expect(data([{ cells }])).toBe(data([cells]));
  });

  it("styles every position in its cells, the empty ones included, and nothing past the last", () => {
    expect(data([{ cells: ["a", null, 1, undefined], style: bold }, ["b"]])).toBe(
      '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>a</t></is></c><c r="B1" s="1"/>' +
        '<c r="C1" s="1"><v>1</v></c><c r="D1" s="1"/></row>' +
        '<row r="2"><c r="A2" t="inlineStr"><is><t>b</t></is></c></row>',
    );
  });

  it("counts a styled empty cell in the used range", () => {
    expect(xml({ name: "S", rows: [{ cells: [null, null, null], style: bold }] })).toContain(
      '<dimension ref="A1:C1"/>',
    );
  });

  it("writes no row for a style over no cells", () => {
    expect(xml({ name: "S", rows: [{ cells: [], style: bold }] })).toContain("<sheetData/>");
  });

  it("lets a cell win over its row, and the row over its column, property by property", () => {
    const styles = createStyleTable();
    const whole: CellStyle = { numberFormat: "0.00", italic: true, bold: true, fill: "#FF0000" };
    const rowOnly: CellStyle = { italic: true, bold: true, fill: "#00FF00" };
    // Registered first, so the indexes below are 1 and 2.
    expect(
      data(
        [
          [
            { value: 1, style: whole },
            { value: 1, style: rowOnly },
          ],
        ],
        {},
        { styles },
      ),
    ).toBe('<row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1" s="2"><v>1</v></c></row>');
    const row = {
      cells: [{ value: 2, style: { fill: "#FF0000" } }, 3, null],
      // A property set to undefined leaves the column's in place.
      style: { bold: true, fill: "#00FF00", italic: undefined } as unknown as CellStyle,
    };
    const columns = [{ style: { numberFormat: "0.00", italic: true, bold: false } }];
    expect(data([row], { columns }, { styles })).toBe(
      '<row r="1"><c r="A1" s="1"><v>2</v></c><c r="B1" s="3"><v>3</v></c><c r="C1" s="3"/></row>',
    );
    // B1 and C1 have no column style: the row's alone, bold on green, a format of its own.
    expect(styles.render()).toContain('<cellXfs count="4">');
  });

  it("merges a row style over a column style once, however many rows carry it", () => {
    const calls: [CellStyle | undefined, CellStyle | undefined][] = [];
    const real = createStyleTable();
    const styles: WorksheetContext["styles"] = {
      indexFor(cell, column, dateDefault) {
        calls.push([cell, column]);
        return real.indexFor(cell, column, dateDefault);
      },
      validate: (style) => real.validate(style),
    };
    const columns = [{ style: { numberFormat: "0.00" } }];
    const rows = [1, 2, 3].map((value) => ({ cells: [value], style: bold }));
    data(rows, { columns }, { styles });
    expect(calls).toHaveLength(3);
    expect(calls[0]![1]).toEqual({ numberFormat: "0.00", bold: true });
    expect(calls[1]![1]).toBe(calls[0]![1]);
    expect(calls[2]![1]).toBe(calls[0]![1]);
  });

  it("takes its height, which wins over the sheet's rowHeights", () => {
    expect(
      data([{ cells: [1], height: 30 }, { cells: [], height: 12.5 }, { cells: [2] }], {
        rowHeights: { 1: 99, 3: 40 },
      }),
    ).toBe(
      '<row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c></row>' +
        '<row r="2" ht="12.5" customHeight="1"/>' +
        '<row r="3" ht="40" customHeight="1"><c r="A3"><v>2</v></c></row>',
    );
  });

  it("measures an auto width with the row's font", () => {
    const widthOf = (rows: Row[]) =>
      /<col min="1" max="1" width="([\d.]+)"/.exec(xml({ name: "S", autoWidth: true, rows }))![1];
    const text = "A header of some length";
    expect(widthOf([{ cells: [text], style: bold }])).toBe(
      widthOf([[{ value: text, style: bold }]]),
    );
    expect(widthOf([{ cells: [text], style: bold }])).not.toBe(widthOf([[text]]));
  });

  it("refuses a row that is neither an array nor an object with cells, at the row", () => {
    for (const row of ["abc", 42, null, {}, { cells: "abc" }, { style: bold }]) {
      expect(() => data([[1], row as unknown as Row])).toThrow(
        expect.objectContaining({
          code: "INVALID_ROW",
          path: "sheets[0].rows[1]",
          message: expect.stringContaining("an array of cells or an object with cells"),
        }),
      );
    }
  });

  it("refuses a height outside what Excel accepts, at the height", () => {
    for (const height of [0, -1, 410, Number.NaN, "20"]) {
      expect(() => data([{ cells: [1], height: height as number }])).toThrow(
        expect.objectContaining({ code: "INVALID_ROW_HEIGHT", path: "sheets[0].rows[0].height" }),
      );
    }
  });

  it("refuses a row style Excel would refuse, at the property", () => {
    expect(() => data([{ cells: [1], style: { fill: "red" } }])).toThrow(
      expect.objectContaining({ code: "INVALID_STYLE", path: "sheets[0].rows[0].style.fill" }),
    );
    expect(() => data([{ cells: [], style: "bold" as unknown as CellStyle }])).toThrow(
      expect.objectContaining({ code: "INVALID_STYLE", path: "sheets[0].rows[0].style" }),
    );
  });

  it("names a cell inside cells in an error, and a cell of an array row as before", () => {
    expect(() => data([{ cells: [1, Number.NaN] }])).toThrow(
      expect.objectContaining({ code: "INVALID_NUMBER", path: "sheets[0].rows[0].cells[1]" }),
    );
    expect(() => data([{ cells: [{ value: 1, style: { fill: "red" } }], style: bold }])).toThrow(
      expect.objectContaining({
        code: "INVALID_STYLE",
        path: "sheets[0].rows[0].cells[0].style.fill",
      }),
    );
    expect(() => data([{ cells: [1] }, [Number.NaN]])).toThrow(
      expect.objectContaining({ code: "INVALID_NUMBER", path: "sheets[0].rows[1][0]" }),
    );
  });

  it("writes no cell for a blank whose row style comes to the default format", () => {
    expect(data([{ cells: [null, 1], style: {} }])).toBe('<row r="1"><c r="B1"><v>1</v></c></row>');
  });

  it("treats a border set to null on the row as not set, like a property set to undefined", () => {
    const styles = createStyleTable();
    const columns = [{ style: { border: "thin" as const } }];
    const style = { bold: true, border: null } as unknown as CellStyle;
    data([[{ value: 1, style: { bold: true, border: "thin" } }]], {}, { styles });
    expect(data([{ cells: [1], style }], { columns }, { styles })).toContain('<c r="A1" s="1">');
  });

  it("lets an error that is not the library's pass through as it is", () => {
    const failure = new TypeError("not ours");
    const styles: WorksheetContext["styles"] = {
      indexFor: () => 0,
      validate: () => {
        throw failure;
      },
    };
    expect(() => data([{ cells: [1], style: bold }], {}, { styles })).toThrow(failure);
  });

  it("holds an object row to the column limit", () => {
    expect(() => data([{ cells: Array.from({ length: 16_385 }, () => null) }])).toThrow(
      expect.objectContaining({ code: "TOO_MANY_COLUMNS", path: "sheets[0].rows[0]" }),
    );
  });
});
