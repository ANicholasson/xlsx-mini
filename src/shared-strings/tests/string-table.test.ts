import { describe, expect, it } from "vite-plus/test";
import { createStringTable } from "../index.js";

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
});
