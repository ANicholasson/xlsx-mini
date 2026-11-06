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

  it("registers one font per distinct bold, italic, size and colour, and one fill and border per value", () => {
    const table = createStyleTable();
    const first = table.indexFor(
      { italic: true, fontSize: 14, fontColor: "#ff0000" },
      undefined,
      NONE,
    );
    expect(
      table.indexFor({ italic: true, fontSize: 14, fontColor: "#FF0000" }, undefined, NONE),
    ).toBe(first);
    table.indexFor({ fill: "#d9e1f2", border: "thin" }, undefined, NONE);
    table.indexFor(
      { fill: "#D9E1F2", border: { top: "thin", right: "thin", bottom: "thin", left: "thin" } },
      undefined,
      NONE,
    );
    table.indexFor({ border: { top: "medium", bottom: "thick" } }, undefined, NONE);
    const xml = table.render();
    expect(xml).toContain('<fonts count="2">');
    expect(xml).toContain(
      '<font><i/><sz val="14"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/></font>',
    );
    expect(xml).toContain('<fills count="3">');
    expect(xml).toContain(
      '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>',
    );
    expect(xml).toContain('<borders count="3">');
    expect(xml).toContain(
      '<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>',
    );
    expect(xml).toContain(
      '<border><left/><right/><top style="medium"><color auto="1"/></top><bottom style="thick"><color auto="1"/></bottom><diagonal/></border>',
    );
    expect(xml).toContain(
      '<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>',
    );
    expect(xml).toContain('<cellXfs count="4">');
  });

  it("aligns an indented cell left unless it says right, and keeps the font size boundaries", () => {
    const table = createStyleTable();
    table.indexFor({ indent: 1 }, undefined, NONE);
    table.indexFor({ indent: 1, align: "left" }, undefined, NONE);
    table.indexFor({ indent: 3, align: "right" }, undefined, NONE);
    table.indexFor({ indent: 0, align: "center" }, undefined, NONE);
    table.indexFor({ fontSize: 1 }, { fontSize: 409 }, NONE);
    table.indexFor({ fontSize: 10.5 }, undefined, NONE);
    const xml = table.render();
    expect(xml).toContain('<cellXfs count="6">');
    expect(xml).toContain('<alignment horizontal="left" indent="1"/>');
    expect(xml).toContain('<alignment horizontal="right" indent="3"/>');
    expect(xml).toContain('<alignment horizontal="center"/>');
    expect(xml).toContain('<sz val="1"/>');
    expect(xml).toContain('<sz val="10.5"/>');
    expect(xml).not.toContain('<sz val="409"/>');
  });

  it("writes an alignment child only for aligned, wrapped or indented formats, with middle as center", () => {
    const table = createStyleTable();
    table.indexFor(
      { align: "right", verticalAlign: "middle", wrap: true, indent: 2 },
      undefined,
      NONE,
    );
    table.indexFor({ wrap: true }, { align: "right" }, NONE);
    table.indexFor({ verticalAlign: "top" }, undefined, NONE);
    table.indexFor({ bold: true }, undefined, NONE);
    const xml = table.render();
    expect(xml).toContain(
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="center" wrapText="1" indent="2"/></xf>',
    );
    expect(xml).toContain('<alignment horizontal="right" wrapText="1"/>');
    expect(xml).toContain('<alignment vertical="top"/>');
    expect(xml).not.toContain('horizontal="left"');
    expect(xml).toContain(
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
    );
    expect(xml.match(/<alignment/g)).toHaveLength(3);
  });

  it("merges every property of a cell style over the column style, and a whole border wins over the column's", () => {
    const table = createStyleTable();
    const column = {
      fill: "#eeeeee",
      border: "thin" as const,
      align: "left" as const,
      numberFormat: "0.0",
    };
    const cell = {
      border: { top: "thick" as const },
      align: "right" as const,
      fontColor: "#000000",
    };
    const index = table.indexFor(cell, column, NONE);
    expect(index).toBe(
      table.indexFor(
        {
          fill: "#eeeeee",
          border: { top: "thick" },
          align: "right",
          fontColor: "#000000",
          numberFormat: "0.0",
        },
        undefined,
        NONE,
      ),
    );
    expect(table.indexFor({ fontSize: 11 }, undefined, NONE)).toBe(0);
    expect(table.indexFor({ border: {} }, undefined, NONE)).toBe(0);
  });

  it("rejects values Excel would refuse, naming the property", () => {
    const table = createStyleTable();
    const bad = (style: object) => {
      try {
        table.indexFor(style as never, undefined, NONE);
      } catch (error) {
        return error as { code: string; path: string };
      }
      throw new Error("expected a throw");
    };
    expect(bad({ fontColor: "red" })).toMatchObject({ code: "INVALID_STYLE", path: "fontColor" });
    expect(bad({ fill: "#fff" })).toMatchObject({ code: "INVALID_STYLE", path: "fill" });
    expect(bad({ fontSize: 0 })).toMatchObject({ code: "INVALID_STYLE", path: "fontSize" });
    expect(bad({ fontSize: 0.5 })).toMatchObject({ code: "INVALID_STYLE", path: "fontSize" });
    expect(bad({ fontSize: 410 })).toMatchObject({ code: "INVALID_STYLE", path: "fontSize" });
    expect(bad({ fontSize: "12" })).toMatchObject({ code: "INVALID_STYLE", path: "fontSize" });
    expect(bad({ bold: "yes" })).toMatchObject({ code: "INVALID_STYLE", path: "bold" });
    expect(bad({ italic: 1 })).toMatchObject({ code: "INVALID_STYLE", path: "italic" });
    expect(bad({ wrap: "false" })).toMatchObject({ code: "INVALID_STYLE", path: "wrap" });
    expect(bad({ align: 10n })).toMatchObject({
      code: "INVALID_STYLE",
      path: "align",
      message: "Alignments are left, center or right, not 10n at align",
    });
    expect(bad({ align: "center", indent: 1 })).toMatchObject({
      code: "INVALID_STYLE",
      path: "indent",
    });
    expect(bad(null as never)).toMatchObject({ code: "INVALID_STYLE", path: "" });
    expect(bad("bold" as never)).toMatchObject({
      code: "INVALID_STYLE",
      path: "",
      message: 'A style is an object, not "bold"',
    });
    expect(bad({ border: "dashed" })).toMatchObject({ code: "INVALID_STYLE", path: "border" });
    expect(bad({ border: { left: "double" } })).toMatchObject({
      code: "INVALID_STYLE",
      path: "border.left",
    });
    expect(bad({ border: 3 })).toMatchObject({ code: "INVALID_STYLE", path: "border" });
    expect(bad({ align: "justify" })).toMatchObject({ code: "INVALID_STYLE", path: "align" });
    expect(bad({ verticalAlign: "center" })).toMatchObject({
      code: "INVALID_STYLE",
      path: "verticalAlign",
    });
    expect(bad({ indent: -1 })).toMatchObject({ code: "INVALID_STYLE", path: "indent" });
    expect(bad({ indent: 251 })).toMatchObject({ code: "INVALID_STYLE", path: "indent" });
    expect(table.render()).toContain('<cellXfs count="1">');
    const validate = (style: object | null) => () => table.validate(style as never);
    expect(validate(null)).toThrow(expect.objectContaining({ code: "INVALID_STYLE", path: "" }));
    expect(validate({ bold: "yes" })).toThrow(expect.objectContaining({ path: "bold" }));
    expect(validate({ italic: 0 })).toThrow(expect.objectContaining({ path: "italic" }));
    expect(validate({ wrap: 1 })).toThrow(expect.objectContaining({ path: "wrap" }));
    expect(validate({ verticalAlign: "center" })).toThrow(
      expect.objectContaining({ path: "verticalAlign" }),
    );
    expect(validate({ align: "center", indent: 2 })).toThrow(
      expect.objectContaining({ path: "indent" }),
    );
    expect(validate({ fontSize: 1 })).not.toThrow();
    expect(validate({ fontSize: 10.5 })).not.toThrow();
    expect(validate({ border: 3 })).toThrow(expect.objectContaining({ path: "border" }));
    expect(validate({ border: null })).not.toThrow();
    expect(table.indexFor({ border: null } as never, undefined, NONE)).toBe(0);
    expect(validate({ border: { right: "none" } })).toThrow(
      expect.objectContaining({ path: "border.right" }),
    );
    expect(validate({ fontColor: "#12345g", fill: "#000000" })).toThrow(
      expect.objectContaining({ path: "fontColor" }),
    );
    expect(
      validate({
        bold: true,
        italic: true,
        fontSize: 9,
        fontColor: "#000000",
        fill: "#ffffff",
        border: "medium",
        align: "left",
        verticalAlign: "top",
        wrap: true,
        indent: 1,
        numberFormat: 1,
      }),
    ).not.toThrow();
  });

  it("resolves a reused style, alone or over a column style, by identity for every date default", () => {
    const table = createStyleTable();
    const column = { numberFormat: "yyyy-mm-dd" };
    const cell = { bold: true };
    const dated = table.indexFor(column, undefined, DATE_TIME);
    const pair = table.indexFor(cell, column, DATE_TIME);
    const pairDate = table.indexFor(cell, column, DATE);
    for (let i = 0; i < 10_000; i++) {
      expect(table.indexFor(undefined, column, DATE_TIME)).toBe(dated);
      expect(table.indexFor(cell, column, DATE_TIME)).toBe(pair);
      expect(table.indexFor(cell, column, DATE)).toBe(pairDate);
    }
    // The explicit format wins over any date default, so the three resolve to two formats.
    expect(dated).not.toBe(pair);
    expect(pairDate).toBe(pair);
    expect(table.indexFor({ bold: true }, { numberFormat: "yyyy-mm-dd" }, DATE)).toBe(pair);
    expect(table.render()).toContain('<cellXfs count="3">');
  });
});
