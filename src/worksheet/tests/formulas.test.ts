import { describe, expect, it } from "vite-plus/test";
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
