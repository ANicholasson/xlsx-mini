import { describe, expect, it } from "vite-plus/test";
import { createStringTable } from "../../shared-strings/index.js";
import { renderWorksheet } from "../index.js";

describe("renderWorksheet", () => {
  it("writes strings as shared-string references and numbers inline", () => {
    const strings = createStringTable();
    const xml = renderWorksheet(
      {
        name: "Sheet1",
        rows: [
          ["Name", "Age"],
          ["Kari", 42.5],
        ],
      },
      strings,
    );
    expect(xml).toContain('<dimension ref="A1:B2"/>');
    expect(xml).toContain(
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>',
    );
    expect(xml).toContain('<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>42.5</v></c></row>');
  });

  it("skips null and undefined cells but keeps their position", () => {
    const xml = renderWorksheet({ name: "S", rows: [[null, undefined, 7]] }, createStringTable());
    expect(xml).toContain('<row r="1"><c r="C1"><v>7</v></c></row>');
    expect(xml).toContain('<dimension ref="A1:C1"/>');
  });

  it("renders an empty sheet with an empty sheetData element", () => {
    const xml = renderWorksheet({ name: "S", rows: [] }, createStringTable());
    expect(xml).toContain("<sheetData/>");
    expect(xml).not.toContain("<dimension");
  });

  it("accepts any iterable of rows", () => {
    function* rows() {
      yield [1];
      yield [2];
    }
    const xml = renderWorksheet({ name: "S", rows: rows() }, createStringTable());
    expect(xml).toContain('<row r="2"><c r="A2"><v>2</v></c></row>');
  });
});
