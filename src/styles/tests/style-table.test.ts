import { describe, expect, it } from "vite-plus/test";
import { createStyleTable } from "../index.js";

describe("style table", () => {
  it("renders only the default cell format when no date style was asked for", () => {
    const xml = createStyleTable().render();
    expect(xml).toContain('<cellXfs count="1">');
    expect(xml).not.toContain("applyNumberFormat");
  });

  it("allocates each date format once, on first use, and returns the same index afterwards", () => {
    const table = createStyleTable();
    expect(table.dateTime()).toBe(1);
    expect(table.dateTime()).toBe(1);
    expect(table.dateOnly()).toBe(2);
    expect(table.dateOnly()).toBe(2);
    const xml = table.render();
    expect(xml).toContain('<cellXfs count="3">');
    expect(xml.indexOf('numFmtId="22"')).toBeLessThan(xml.indexOf('numFmtId="14"'));
  });
});
