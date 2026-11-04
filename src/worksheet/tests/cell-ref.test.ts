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
});
