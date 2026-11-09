import { describe, expect, it } from "vite-plus/test";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import { type DataValidation, type Sheet, renderValidations, renderWorksheet } from "../index.js";

const names = ["Orders", "Lists", "Per O'Brien"];
const one = (rule: Record<string, unknown>) =>
  renderValidations([rule as unknown as DataValidation], 0, names);
const failing = (rule: Record<string, unknown>) => () => one(rule);

describe("renderValidations", () => {
  it("writes nothing without rules", () => {
    expect(renderValidations(undefined, 0, names)).toBe("");
    expect(renderValidations([], 0, names)).toBe("");
  });

  it("writes a literal list as Excel's quoted source, quotes doubled and escaped", () => {
    expect(one({ range: "B2:B100", list: ["Open", 'Say "hi"', "On hold"] })).toBe(
      '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="B2:B100"><formula1>&quot;Open,Say &quot;&quot;hi&quot;&quot;,On hold&quot;</formula1></dataValidation></dataValidations>',
    );
  });

  it("writes a list read from a range as an absolute reference, with the sheet quoted when named", () => {
    expect(one({ range: "C2", list: "A2:A4" })).toContain("<formula1>$A$2:$A$4</formula1>");
    expect(one({ range: "C2", list: "Lists!A2:A4" })).toContain(
      "<formula1>&apos;Lists&apos;!$A$2:$A$4</formula1>",
    );
    expect(one({ range: "C2", list: "'Per O''Brien'!B:B" })).toContain(
      "<formula1>&apos;Per O&apos;&apos;Brien&apos;!$B$1:$B$1048576</formula1>",
    );
    expect(one({ range: "C2", list: "lists!A2" })).toContain("&apos;lists&apos;!$A$2");
  });

  it("writes number bounds as whole or decimal with the operator the bounds imply", () => {
    expect(one({ range: "D2:D9", number: { min: 1, max: 1000, integer: true } })).toContain(
      '<dataValidation type="whole" operator="between" allowBlank="1" sqref="D2:D9"><formula1>1</formula1><formula2>1000</formula2></dataValidation>',
    );
    expect(one({ range: "D2", number: { min: 0.5 } })).toContain(
      'type="decimal" operator="greaterThanOrEqual" allowBlank="1" sqref="D2"><formula1>0.5</formula1></dataValidation>',
    );
    expect(one({ range: "D2", number: { max: -1, integer: false } })).toContain(
      'type="decimal" operator="lessThanOrEqual" allowBlank="1" sqref="D2"><formula1>-1</formula1></dataValidation>',
    );
  });

  it("writes date bounds as UTC serials", () => {
    expect(
      one({
        range: "E2:E9",
        date: { min: new Date(Date.UTC(2026, 0, 1)), max: new Date(Date.UTC(2026, 11, 31, 12)) },
      }),
    ).toContain(
      '<dataValidation type="date" operator="between" allowBlank="1" sqref="E2:E9"><formula1>46023</formula1><formula2>46387.5</formula2></dataValidation>',
    );
    expect(one({ range: "E2", date: { max: new Date(Date.UTC(2026, 0, 1)) } })).toContain(
      'operator="lessThanOrEqual"',
    );
  });

  it("writes whole columns as the rows Excel stores, and several rules with a count", () => {
    const xml = renderValidations(
      [
        { range: "F:F", number: { max: 1 } },
        { range: "A1", list: ["x"] },
      ],
      0,
      names,
    );
    expect(xml).toContain('<dataValidations count="2">');
    expect(xml).toContain('sqref="F1:F1048576"');
    expect(xml).toContain('sqref="A1"');
  });

  it("writes the input and error messages, the error style and allowBlank", () => {
    expect(
      one({
        range: "B2",
        list: ["a"],
        allowBlank: false,
        input: { title: "Status", message: "Pick <one>" },
        error: { title: "No", message: "Not & allowed", style: "warning" },
      }),
    ).toContain(
      '<dataValidation type="list" showInputMessage="1" promptTitle="Status" prompt="Pick &lt;one&gt;" showErrorMessage="1" errorStyle="warning" errorTitle="No" error="Not &amp; allowed" sqref="B2">',
    );
    expect(one({ range: "B2", list: ["a"], error: { message: "m", style: "stop" } })).toContain(
      'allowBlank="1" showErrorMessage="1" error="m" sqref',
    );
    expect(one({ range: "B2", list: ["a"], input: { title: "", message: "m" } })).toContain(
      'showInputMessage="1" prompt="m"',
    );
  });

  it("reports what is wrong with a rule, at the property", () => {
    const cases: readonly [Record<string, unknown>, string, string][] = [
      [{ range: "B2:B1", list: ["a"] }, ".range", "top-left cell"],
      [{ range: "XFE1", list: ["a"] }, ".range", "outside the sheet"],
      [{ range: "b2", list: ["a"] }, ".range", 'such as "B2:B100", not "b2"'],
      [{ range: "B2" }, "", "exactly one of list, number and date"],
      [{ range: "B2", list: ["a"], number: { min: 1 } }, "", "exactly one"],
      [{ range: "B2", list: [] }, ".list", "non-empty array"],
      [{ range: "B2", list: ["a", 1] }, ".list[1]", "strings, not 1"],
      [{ range: "B2", list: ["a,b"] }, ".list[0]", "comma"],
      [{ range: "B2", list: ["a".repeat(254)] }, ".list", "at most 255 characters"],
      [{ range: "B2", list: "Open,Closed" }, ".list", 'such as "Lists!A2:A9"'],
      [{ range: "B2", list: "Nope!A1" }, ".list", 'no sheet named "Nope"'],
      [{ range: "B2", list: "A0" }, ".list", '"A0" is not a cell or range'],
      [{ range: "B2", list: "Lists!A01:A02" }, ".list", '"A01:A02" is not a cell or range'],
      [{ range: "B2", list: "A9999999" }, ".list", "outside the sheet"],
      [{ range: "B2", number: 5 }, ".number", "object with min and max"],
      [{ range: "B2", number: {} }, ".number", "min, a max or both"],
      [{ range: "B2", number: { min: 5, max: 1 } }, ".number", "max 1 is below min 5"],
      [{ range: "B2", number: { min: Number.NaN } }, ".number.min", "finite number"],
      [{ range: "B2", number: { max: "9" } }, ".number.max", "finite number"],
      [{ range: "B2", number: { min: 1, integer: "yes" } }, ".number.integer", "true or false"],
      [{ range: "B2", date: { min: "2026-01-01" } }, ".date.min", "valid Date"],
      [{ range: "B2", date: { max: new Date(Number.NaN) } }, ".date.max", "valid Date"],
      [{ range: "B2", date: { min: new Date(Date.UTC(1800, 0, 1)) } }, ".date.min", "1900-01-01"],
      [{ range: "B2", list: ["a"], allowBlank: 1 }, ".allowBlank", "true or false"],
      [{ range: "B2", list: ["a"], input: "hi" }, ".input", "object with a message"],
      [{ range: "B2", list: ["a"], input: { message: "" } }, ".input.message", "1 to 255"],
      [
        { range: "B2", list: ["a"], input: { title: "t".repeat(33), message: "m" } },
        ".input.title",
        "at most 32",
      ],
      [
        { range: "B2", list: ["a"], error: { message: "m", style: "loud" } },
        ".error.style",
        "stop, warning or information",
      ],
      [
        { range: "B2", list: ["a"], error: { message: "m".repeat(256) } },
        ".error.message",
        "1 to 255",
      ],
    ];
    for (const [rule, suffix, message] of cases) {
      expect(failing(rule), JSON.stringify(rule)).toThrow(
        expect.objectContaining({
          code: "INVALID_VALIDATION",
          path: `sheets[0].validations[0]${suffix}`,
          message: expect.stringContaining(message),
        }),
      );
    }
    expect(() => renderValidations([null as never], 2, names)).toThrow(
      expect.objectContaining({ code: "INVALID_VALIDATION", path: "sheets[2].validations[0]" }),
    );
  });

  it("is written after the merges and before the hyperlinks in the sheet part", () => {
    const sheet: Sheet = {
      name: "Orders",
      merges: ["A1:B1"],
      validations: [{ range: "B2", list: ["a"] }],
      rows: [["t"], ["x", { value: "link", href: "https://example.com" }]],
    };
    const xml = new TextDecoder().decode(
      renderWorksheet(sheet, {
        strings: createStringTable(),
        styles: createStyleTable(),
        coerce: undefined,
        sheetIndex: 0,
        sheetNames: names,
      }).data,
    );
    expect(xml).toMatch(/<\/mergeCells><dataValidations .*<\/dataValidations><hyperlinks>/);
  });
});
