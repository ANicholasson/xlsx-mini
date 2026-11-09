import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { absolute, boxProblem, columnLetter, parseReference, quoteSheetName } from "../index.js";

describe("columnLetter", () => {
  it("converts zero-based column indexes to Excel letters", () => {
    expect(columnLetter(0)).toBe("A");
    expect(columnLetter(25)).toBe("Z");
    expect(columnLetter(26)).toBe("AA");
    expect(columnLetter(51)).toBe("AZ");
    expect(columnLetter(52)).toBe("BA");
    expect(columnLetter(701)).toBe("ZZ");
    expect(columnLetter(702)).toBe("AAA");
    expect(columnLetter(16383)).toBe("XFD");
  });

  it("round-trips through an independent letters-to-index conversion", () => {
    const toIndex = (letters: string): number => {
      let index = 0;
      for (const letter of letters) index = index * 26 + (letter.charCodeAt(0) - 64);
      return index - 1;
    };
    fc.assert(
      fc.property(fc.integer({ min: 0, max: 16383 }), (index) => {
        const letters = columnLetter(index);
        expect(letters).toMatch(/^[A-Z]{1,3}$/);
        expect(toIndex(letters)).toBe(index);
      }),
    );
  });
});

describe("parseReference", () => {
  it("parses a cell, a range and whole columns, expanding columns to every row", () => {
    expect(parseReference("B2")).toEqual({ ref: "B2", top: 2, left: 1, bottom: 2, right: 1 });
    expect(parseReference("A1:D4")).toEqual({ ref: "A1:D4", top: 1, left: 0, bottom: 4, right: 3 });
    expect(parseReference("B:C")).toEqual({
      ref: "B1:C1048576",
      top: 1,
      left: 1,
      bottom: 1_048_576,
      right: 2,
    });
    expect(parseReference("XFD1048576")?.right).toBe(16_383);
  });

  it("returns undefined for anything else", () => {
    for (const ref of ["", "b2", "$A$1", "A", "1", "A1:B", "A0", "Sheet!A1", 42, null, "A1 B1"]) {
      expect(parseReference(ref), String(ref)).toBeUndefined();
    }
  });

  it("names a box outside the sheet or written backwards", () => {
    expect(boxProblem(parseReference("A1:D4")!)).toBeUndefined();
    expect(boxProblem(parseReference("XFE1")!)).toBe("Range XFE1 lies outside the sheet");
    expect(boxProblem(parseReference("A1048577")!)).toBe("Range A1048577 lies outside the sheet");
    expect(boxProblem(parseReference("D4:A1")!)).toBe(
      "Range D4:A1 must run from its top-left cell to its bottom-right cell",
    );
    expect(boxProblem(parseReference("C:B")!)).toContain("must run from its top-left cell");
  });

  it("makes references absolute and quotes sheet names", () => {
    expect(absolute("A1:C40")).toBe("$A$1:$C$40");
    expect(quoteSheetName("Per O'Brien")).toBe("'Per O''Brien'");
  });
});
