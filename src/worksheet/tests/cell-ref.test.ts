import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { columnLetter } from "../index.js";

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
