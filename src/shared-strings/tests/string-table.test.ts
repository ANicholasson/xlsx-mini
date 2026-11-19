import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import {
  textsToReadBack,
  keptByXml,
  readTextAsExcel,
  readTextAsNumbers,
} from "../../../test/helpers/text.js";
import { createStringTable } from "../index.js";

function readBack(text: string): [string, string] {
  const table = createStringTable();
  table.intern(text);
  const item = /<si><t(?: xml:space="preserve")?>([^<]*)<\/t><\/si>/.exec(table.render())?.[1];
  if (item === undefined) throw new Error("no string item rendered");
  return [readTextAsExcel(item), readTextAsNumbers(item)];
}

describe("string table", () => {
  it("interns repeated strings to one index and counts every use", () => {
    const table = createStringTable();
    expect(table.intern("a")).toBe(0);
    expect(table.intern("b")).toBe(1);
    expect(table.intern("a")).toBe(0);
    expect(table.render()).toContain('count="3" uniqueCount="2"');
  });

  it("escapes text and preserves significant whitespace", () => {
    const table = createStringTable();
    table.intern("Tom & Jerry");
    table.intern(" padded ");
    const xml = table.render();
    expect(xml).toContain("<si><t>Tom &amp; Jerry</t></si>");
    expect(xml).toContain('<si><t xml:space="preserve"> padded </t></si>');
  });

  it.each(textsToReadBack)("writes %j so that it reads back as written", (text) => {
    expect(readBack(text)).toEqual([text, text]);
  });

  it("looks for an escape in the text that XML keeps, not in the text given", () => {
    const kept = "_x0041_";
    expect(keptByXml("_x\u00000041_")).toBe(kept);
    expect(readBack("_x\u00000041_")).toEqual([kept, kept]);
  });

  it("writes any text so that it reads back as written, less what XML cannot carry", () => {
    const piece = fc.oneof(
      fc.constantFrom("_", "x", "_x", "0041", "00e6", "005F", "000D", "ZZZZ", "\r", "\n"),
      fc.constantFrom("_x0041_", "_x005F_", "x0042_", "_x00e6"),
      fc.constantFrom("&", "<", "\u0000", "\uD83D", "\uDE00"),
      fc.string({ unit: "binary", maxLength: 3 }),
    );
    fc.assert(
      fc.property(
        fc.array(piece, { maxLength: 12 }).map((pieces) => pieces.join("")),
        (text) => {
          const kept = keptByXml(text);
          expect(readBack(text)).toEqual([kept, kept]);
        },
      ),
    );
  });
});
