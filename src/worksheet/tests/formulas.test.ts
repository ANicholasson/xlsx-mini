import { describe, expect, it } from "vite-plus/test";
import { createMediaTable } from "../../drawing/index.js";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import {
  type RenderedWorksheet,
  type Sheet,
  type WorksheetContext,
  renderWorksheet,
} from "../index.js";

function render(sheet: Sheet, overrides: Partial<WorksheetContext> = {}): RenderedWorksheet {
  return renderWorksheet(sheet, {
    strings: createStringTable(),
    media: createMediaTable(),
    styles: createStyleTable(),
    coerce: undefined,
    sheetIndex: 0,
    sheetNames: ["S"],
    ...overrides,
  });
}

const xml = (sheet: Sheet, overrides?: Partial<WorksheetContext>) =>
  new TextDecoder().decode(render(sheet, overrides).data);

describe("formula cells", () => {
  it("writes the formula escaped inside <f>, without a value", () => {
    const sheet = xml({ name: "S", rows: [[{ formula: 'IF(A2>1,"a & b","c")' }]] });
    expect(sheet).toContain('<c r="A1"><f>IF(A2&gt;1,&quot;a &amp; b&quot;,&quot;c&quot;)</f></c>');
  });

  it("writes a cached value typed as the result: number, text as str, boolean", () => {
    const sheet = xml({
      name: "S",
      rows: [
        [
          { formula: "SUM(B2:B4)", value: 7 },
          { formula: "A1&B1", value: "7 <ok>" },
          { formula: "A1>1", value: true },
          { formula: "A1>1", value: null },
          { formula: "A1<1", value: false },
        ],
      ],
    });
    expect(sheet).toContain('<c r="A1"><f>SUM(B2:B4)</f><v>7</v></c>');
    expect(sheet).toContain('<c r="B1" t="str"><f>A1&amp;B1</f><v>7 &lt;ok&gt;</v></c>');
    expect(sheet).toContain('<c r="C1" t="b"><f>A1&gt;1</f><v>1</v></c>');
    expect(sheet).toContain('<c r="D1"><f>A1&gt;1</f></c>');
    expect(sheet).toContain('<c r="E1" t="b"><f>A1&lt;1</f><v>0</v></c>');
    expect(sheet).not.toContain('t="s"');
  });

  it("styles a formula cell from its own style and its column's", () => {
    const sheet = xml({
      name: "S",
      columns: [{ style: { numberFormat: "0.00" } }],
      rows: [[{ formula: "1/3" }], [{ formula: "1/3", style: { bold: true } }]],
    });
    expect(sheet).toContain('<c r="A1" s="1"><f>1/3</f></c>');
    expect(sheet).toContain('<c r="A2" s="2"><f>1/3</f></c>');
  });

  it("reports whether the sheet holds a formula, for the workbook's calculation flag", () => {
    expect(render({ name: "S", rows: [["=SUM(A1)"]] }).hasFormulas).toBe(false);
    expect(render({ name: "S", rows: [[{ formula: "SUM(A1)" }]] }).hasFormulas).toBe(true);
  });

  it("keeps text that starts with = as text, never a formula", () => {
    const strings = createStringTable();
    const sheet = xml({ name: "S", rows: [["=SUM(A1:A2)"]] }, { strings });
    expect(sheet).toContain('<c r="A1" t="s"><v>0</v></c>');
    expect(strings.render()).toContain("<t>=SUM(A1:A2)</t>");
  });

  it("refuses a leading =, an empty formula, a long one and a non-string, at the formula", () => {
    const cases: readonly [unknown, string][] = [
      ["=SUM(A1)", "without the leading ="],
      ["", "1 to 8192 characters"],
      ["A".repeat(8193), "1 to 8192 characters"],
      [42, "1 to 8192 characters"],
    ];
    for (const [formula, message] of cases) {
      expect(() => xml({ name: "S", rows: [[{ formula }]] } as Sheet)).toThrow(
        expect.objectContaining({
          code: "INVALID_FORMULA",
          path: "sheets[0].rows[0][0].formula",
          message: expect.stringContaining(message),
        }),
      );
    }
  });

  it("checks a cached value as it checks a cell value, at the value", () => {
    expect(() => xml({ name: "S", rows: [[{ formula: "A2", value: Number.NaN }]] })).toThrow(
      expect.objectContaining({ code: "INVALID_NUMBER", path: "sheets[0].rows[0][0].value" }),
    );
    expect(() =>
      xml({ name: "S", rows: [[{ formula: "A2", value: "x".repeat(32_768) }]] }),
    ).toThrow(
      expect.objectContaining({ code: "CELL_TEXT_TOO_LONG", path: "sheets[0].rows[0][0].value" }),
    );
    expect(() => xml({ name: "S", rows: [[{ formula: "TODAY()", value: new Date() }]] })).toThrow(
      expect.objectContaining({
        code: "INVALID_FORMULA",
        path: "sheets[0].rows[0][0].value",
        message: expect.stringContaining("text, a number or a boolean, not Date"),
      }),
    );
  });

  it("refuses a formula cell that also carries a link, rather than dropping the link", () => {
    for (const extra of [{ href: "https://example.com" }, { cell: "A1" }, { sheet: "S" }]) {
      expect(() => xml({ name: "S", rows: [[{ formula: "1", ...extra } as never]] })).toThrow(
        expect.objectContaining({
          code: "INVALID_FORMULA",
          path: "sheets[0].rows[0][0]",
          message: expect.stringContaining("cannot also be a link"),
        }),
      );
    }
  });

  it("reports a bad style on a formula cell at the cell's style", () => {
    expect(() =>
      xml({ name: "S", rows: [[{ formula: "1", style: { fontColor: "red" } }]] }),
    ).toThrow(
      expect.objectContaining({
        code: "INVALID_STYLE",
        path: "sheets[0].rows[0][0].style.fontColor",
      }),
    );
  });
});

describe("functions newer than Excel 2007", () => {
  /** The formula of a one-cell sheet as the file stores it, unescaped. */
  const stored = (formula: string) =>
    /<f>(.*)<\/f>/
      .exec(xml({ name: "S", rows: [[{ formula }]] }))![1]!
      .replaceAll("&quot;", '"')
      .replaceAll("&gt;", ">")
      .replaceAll("&lt;", "<")
      .replaceAll("&apos;", "'")
      .replaceAll("&amp;", "&");

  it("holds a formula to its length as stored, prefixes included", () => {
    // 481 calls take 8 176 characters once each has its prefix; Excel repairs a file at 8 193.
    const calls = Array.from({ length: 481 }, () => "IFNA(A2,0)").join("+");
    expect(stored(`${calls}${"+1".repeat(8)}`)).toHaveLength(8192);
    expect(() => xml({ name: "S", rows: [[{ formula: `${calls}${"+1".repeat(7)}+11` }]] })).toThrow(
      expect.objectContaining({
        code: "INVALID_FORMULA",
        path: "sheets[0].rows[0][0].formula",
        message: expect.stringContaining("8193 characters as stored"),
      }),
    );
  });

  it("stores them under their prefixed names, so a formula works as typed", () => {
    expect(stored("XLOOKUP(A1,B1:B9,C1:C9)")).toBe("_xlfn.XLOOKUP(A1,B1:B9,C1:C9)");
    expect(stored("STDEV.S(A1:A9)")).toBe("_xlfn.STDEV.S(A1:A9)");
    expect(stored('TEXTJOIN("-",TRUE,A1:A3)')).toBe('_xlfn.TEXTJOIN("-",TRUE,A1:A3)');
    expect(stored("T.DIST.2T(A1,4)")).toBe("_xlfn.T.DIST.2T(A1,4)");
  });

  it("stores FILTER and SORT under their longer prefix", () => {
    expect(stored("SUM(FILTER(A1:A9,A1:A9>4))")).toBe("SUM(_xlfn._xlws.FILTER(A1:A9,A1:A9>4))");
    expect(stored("SORT(A1:A9)")).toBe("_xlfn._xlws.SORT(A1:A9)");
  });

  it("prefixes every call in a formula, nested ones included, in the name's own capitals", () => {
    expect(stored('ifs(A1>5,"big",TRUE,concat(B1:B3))&maxifs(A1:A9,A1:A9,"<6")')).toBe(
      '_xlfn.IFS(A1>5,"big",TRUE,_xlfn.CONCAT(B1:B3))&_xlfn.MAXIFS(A1:A9,A1:A9,"<6")',
    );
  });

  it("leaves older functions, names without a call and formulas without one as they are", () => {
    expect(stored("SUM(A1:A9)+LOG10(B1)")).toBe("SUM(A1:A9)+LOG10(B1)");
    expect(stored("A1*B1")).toBe("A1*B1");
    expect(stored("(A1+XOR)*2")).toBe("(A1+XOR)*2");
    expect(stored("Sheet2!A1+SUM(Sheet2!B1:B2)")).toBe("Sheet2!A1+SUM(Sheet2!B1:B2)");
  });

  it("leaves a name that already carries its prefix alone", () => {
    expect(stored("_xlfn.XLOOKUP(A1,B1:B9,C1:C9)")).toBe("_xlfn.XLOOKUP(A1,B1:B9,C1:C9)");
    expect(stored("_xlfn._xlws.FILTER(A1:A9,A1:A9>4)")).toBe("_xlfn._xlws.FILTER(A1:A9,A1:A9>4)");
  });

  it("passes over text in quotes and sheet names in apostrophes", () => {
    expect(stored('IF(A1="XLOOKUP(","IFS(""x"")",CONCAT(B1:B2))')).toBe(
      'IF(A1="XLOOKUP(","IFS(""x"")",_xlfn.CONCAT(B1:B2))',
    );
    expect(stored("SUM('IFS(a''s)'!A1:A9)+DAYS(B1,C1)")).toBe(
      "SUM('IFS(a''s)'!A1:A9)+_xlfn.DAYS(B1,C1)",
    );
    // A quote that never closes takes the rest of the formula with it.
    expect(stored('SUM(A1)&"IFS(')).toBe('SUM(A1)&"IFS(');
  });

  it("writes a formula that calls LET or LAMBDA exactly as given", () => {
    // Their parameters are stored under a prefix as well, which only the caller can write.
    expect(stored("LET(x,XLOOKUP(A1,B1:B9,C1:C9),x*2)")).toBe("LET(x,XLOOKUP(A1,B1:B9,C1:C9),x*2)");
    expect(stored("_xlfn.LET(_xlpm.x,CONCAT(A1:A2),_xlpm.x)")).toBe(
      "_xlfn.LET(_xlpm.x,CONCAT(A1:A2),_xlpm.x)",
    );
    expect(stored("MAP(A1:A3,LAMBDA(v,v*2))")).toBe("MAP(A1:A3,LAMBDA(v,v*2))");
  });
});
