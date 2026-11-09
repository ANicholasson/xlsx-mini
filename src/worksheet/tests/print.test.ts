import { describe, expect, it } from "vite-plus/test";
import { createStringTable } from "../../shared-strings/index.js";
import { createStyleTable } from "../../styles/index.js";
import { type PrintSetup, type Sheet, renderPrint, renderWorksheet } from "../index.js";

const DEFAULT_MARGINS =
  '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

describe("renderPrint", () => {
  it("renders nothing without a setup", () => {
    expect(renderPrint(undefined, 0)).toEqual({ sheetPr: "", tail: "", printTitles: undefined });
  });

  it("writes Excel's default margins once any setup is given, and no pageSetup without one", () => {
    expect(renderPrint({}, 0)).toEqual({
      sheetPr: "",
      tail: DEFAULT_MARGINS,
      printTitles: undefined,
    });
    expect(renderPrint({ margins: { top: 1, footer: 0 } }, 0).tail).toBe(
      '<pageMargins left="0.7" right="0.7" top="1" bottom="0.75" header="0.3" footer="0"/>',
    );
  });

  it("writes orientation, paper size and fit, with fitToPage in sheetPr", () => {
    const print = renderPrint(
      { orientation: "landscape", paperSize: "letter", fitTo: { width: 1 } },
      0,
    );
    expect(print.sheetPr).toBe('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(print.tail).toBe(
      `${DEFAULT_MARGINS}<pageSetup paperSize="1" orientation="landscape" fitToWidth="1" fitToHeight="0"/>`,
    );
    expect(renderPrint({ fitTo: { height: 2 } }, 0).tail).toContain(
      '<pageSetup fitToWidth="0" fitToHeight="2"/>',
    );
    expect(renderPrint({ paperSize: "A3" }, 0).tail).toContain('<pageSetup paperSize="8"/>');
    expect(renderPrint({ orientation: "portrait" }, 0).sheetPr).toBe("");
  });

  it("writes the header and footer escaped, and the repeated rows as print titles", () => {
    const print = renderPrint({ header: "&L<Q&&A>&RPage &P", repeatRows: 2 }, 0);
    expect(print.tail).toBe(
      `${DEFAULT_MARGINS}<headerFooter><oddHeader>&amp;L&lt;Q&amp;&amp;A&gt;&amp;RPage &amp;P</oddHeader></headerFooter>`,
    );
    expect(print.printTitles).toBe("$1:$2");
    expect(renderPrint({ footer: "&A" }, 0).tail).toContain(
      "<headerFooter><oddFooter>&amp;A</oddFooter></headerFooter>",
    );
  });

  it("reports what is wrong with a setup, at the property", () => {
    const cases: readonly [unknown, string, string][] = [
      [42, "", "object, not 42"],
      [{ orientation: "sideways" }, ".orientation", "portrait or landscape"],
      [{ paperSize: "B5" }, ".paperSize", "letter, tabloid, legal, A3, A4, A5, not"],
      [{ paperSize: "toString" }, ".paperSize", "not"],
      [{ fitTo: 1 }, ".fitTo", "width and height in pages"],
      [{ fitTo: {} }, ".fitTo", "a width, a height or both"],
      [{ fitTo: { width: 0 } }, ".fitTo.width", "1 to 32767"],
      [{ fitTo: { height: 1.5 } }, ".fitTo.height", "1 to 32767"],
      [{ margins: "1cm" }, ".margins", "sides in inches"],
      [{ margins: { left: -1 } }, ".margins.left", "0 to 49 inches"],
      [{ margins: { header: "0.3" } }, ".margins.header", "0 to 49 inches"],
      [{ header: "x".repeat(256) }, ".header", "at most 255"],
      [{ footer: 7 }, ".footer", "at most 255"],
      [{ repeatRows: 0 }, ".repeatRows", "1 to 1048576"],
      [{ repeatRows: 1.5 }, ".repeatRows", "1 to 1048576"],
    ];
    for (const [print, suffix, message] of cases) {
      expect(() => renderPrint(print as PrintSetup, 3), JSON.stringify(print)).toThrow(
        expect.objectContaining({
          code: "INVALID_PRINT_SETUP",
          path: `sheets[3].print${suffix}`,
          message: expect.stringContaining(message),
        }),
      );
    }
  });

  it("places sheetPr first and the print elements last in the sheet part, and reports the titles", () => {
    const sheet: Sheet = {
      name: "S",
      freeze: { rows: 1 },
      merges: ["A1:B1"],
      print: { fitTo: { width: 1 }, footer: "&P", repeatRows: 1 },
      rows: [["a", { value: "b", href: "https://example.com" }]],
    };
    const rendered = renderWorksheet(sheet, {
      strings: createStringTable(),
      styles: createStyleTable(),
      coerce: undefined,
      sheetIndex: 0,
      sheetNames: ["S"],
    });
    const xml = new TextDecoder().decode(rendered.data);
    expect(xml).toMatch(
      /<worksheet [^>]*><sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr><dimension ref="A1:B1"\/><sheetViews>/,
    );
    expect(xml).toMatch(
      /<\/mergeCells><hyperlinks>.*<\/hyperlinks><pageMargins [^>]*\/><pageSetup fitToWidth="1" fitToHeight="0"\/><headerFooter><oddFooter>&amp;P<\/oddFooter><\/headerFooter><\/worksheet>$/,
    );
    expect(rendered.printTitles).toBe("$1:$1");
  });
});
