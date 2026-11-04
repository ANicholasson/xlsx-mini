import { describe, expect, it } from "vite-plus/test";
import { createStyleTable } from "../index.js";

const NONE = 0;
const DATE = 14;
const DATE_TIME = 22;

describe("style table", () => {
  it("renders only the default cell format and one font when nothing was asked for", () => {
    const xml = createStyleTable().render();
    expect(xml).toContain('<fonts count="1">');
    expect(xml).toContain('<cellXfs count="1">');
    expect(xml).not.toContain("<numFmts");
    expect(xml).not.toContain("applyNumberFormat");
  });

  it("hands out index 0 for no style and a date default only when there is no explicit format", () => {
    const table = createStyleTable();
    expect(table.indexFor(undefined, undefined, NONE)).toBe(0);
    expect(table.indexFor(undefined, undefined, DATE_TIME)).toBe(1);
    expect(table.indexFor(undefined, undefined, DATE_TIME)).toBe(1);
    expect(table.indexFor(undefined, undefined, DATE)).toBe(2);
    expect(table.indexFor({ numberFormat: "yyyy-mm-dd" }, undefined, DATE_TIME)).toBe(3);
    const xml = table.render();
    expect(xml).toContain('<cellXfs count="4">');
    expect(xml).toContain('<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>');
  });

  it("deduplicates structurally equal styles and takes the identity fast path for a reused object", () => {
    const table = createStyleTable();
    const money = { numberFormat: '#,##0.00 "kr"' };
    const first = table.indexFor(money, undefined, NONE);
    for (let i = 0; i < 10_000; i++) expect(table.indexFor(money, undefined, NONE)).toBe(first);
    expect(table.indexFor({ numberFormat: '#,##0.00 "kr"' }, undefined, NONE)).toBe(first);
    expect(table.indexFor(undefined, money, NONE)).toBe(first);
    expect(table.indexFor({ bold: true }, undefined, NONE)).not.toBe(first);
    const xml = table.render();
    expect(xml).toContain('<numFmts count="1">');
    expect(xml).toContain('<numFmt numFmtId="164" formatCode="#,##0.00 &quot;kr&quot;"/>');
    expect(xml).toContain('<cellXfs count="3">');
  });

  it("merges a cell style over a column style property by property", () => {
    const table = createStyleTable();
    const column = { numberFormat: "0.0%" };
    const header = { bold: true };
    const index = table.indexFor(header, column, NONE);
    expect(index).toBe(table.indexFor({ bold: true, numberFormat: "0.0%" }, undefined, NONE));
    expect(table.indexFor({ numberFormat: "0%" }, column, NONE)).not.toBe(index);
    expect(table.indexFor({ bold: false }, { bold: true }, NONE)).toBe(0);
  });

  it("adds one bold font only when bold is used, and marks what each cell format applies", () => {
    const table = createStyleTable();
    table.indexFor({ bold: true }, undefined, NONE);
    table.indexFor({ bold: true, numberFormat: 10 }, undefined, NONE);
    const xml = table.render();
    expect(xml).toContain('<fonts count="2">');
    expect(xml).toContain("<b/>");
    expect(xml).toContain(
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    );
    expect(xml).toContain(
      '<xf numFmtId="10" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>',
    );
    expect(xml.indexOf("<numFmts")).toBe(-1);
  });

  it("puts custom number formats before fonts, as the schema requires", () => {
    const table = createStyleTable();
    table.indexFor({ numberFormat: "0.00" }, undefined, NONE);
    const xml = table.render();
    expect(xml.indexOf("<numFmts")).toBeLessThan(xml.indexOf("<fonts"));
  });

  it("rejects number formats Excel would refuse", () => {
    const table = createStyleTable();
    for (const bad of ["", "x".repeat(256), -1, 50, 164, 1.5, Number.NaN]) {
      expect(() => table.indexFor({ numberFormat: bad }, undefined, NONE)).toThrow(
        expect.objectContaining({ code: "INVALID_NUMBER_FORMAT" }),
      );
    }
  });
});
