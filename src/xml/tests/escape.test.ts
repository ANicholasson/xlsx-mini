import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { escapeXml, needsSpacePreserve } from "../index.js";

describe("escapeXml", () => {
  it("escapes the five XML special characters", () => {
    expect(escapeXml(`Tom & Jerry <"quoted"> 'single'`)).toBe(
      "Tom &amp; Jerry &lt;&quot;quoted&quot;&gt; &apos;single&apos;",
    );
  });

  it("returns the same string when nothing needs escaping", () => {
    const text = "Ærlig norsk tekst – med dash og emoji 🎉";
    expect(escapeXml(text)).toBe(text);
  });

  it("handles an empty string", () => {
    expect(escapeXml("")).toBe("");
  });

  it("drops characters XML 1.0 cannot carry but keeps tab, newline and carriage return", () => {
    expect(escapeXml("a\u0000b\u0007c\u001fd")).toBe("abcd");
    expect(escapeXml("line1\nline2\ttab\r")).toBe("line1\nline2\ttab\r");
    expect(escapeXml("x￾y￿z")).toBe("xyz");
  });

  it("drops lone surrogates and keeps valid pairs", () => {
    expect(escapeXml("a\uD83Cb")).toBe("ab");
    expect(escapeXml("a\uDF89b")).toBe("ab");
    expect(escapeXml("ab\uD83C")).toBe("ab");
    expect(escapeXml("a🎉b")).toBe("a🎉b");
  });

  it("agrees with an independent strip-then-escape oracle on arbitrary input", () => {
    const oracle = (text: string): string => {
      let kept = "";
      for (const character of text) {
        const point = character.codePointAt(0) ?? 0;
        const valid =
          point === 0x9 ||
          point === 0xa ||
          point === 0xd ||
          (point >= 0x20 && point <= 0xd7ff) ||
          (point >= 0xe000 && point <= 0xfffd) ||
          (point >= 0x10000 && point <= 0x10ffff);
        if (valid) kept += character;
      }
      return kept
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&apos;");
    };
    const anyCodeUnits = fc
      .array(fc.integer({ min: 0, max: 0xffff }), { maxLength: 40 })
      .map((units) => String.fromCharCode(...units));
    fc.assert(
      fc.property(fc.oneof(fc.string({ unit: "binary" }), anyCodeUnits), (text) => {
        expect(escapeXml(text)).toBe(oracle(text));
      }),
    );
  });
});

describe("needsSpacePreserve", () => {
  it("is true only for leading or trailing whitespace", () => {
    expect(needsSpacePreserve(" padded")).toBe(true);
    expect(needsSpacePreserve("padded\t")).toBe(true);
    expect(needsSpacePreserve("in the middle")).toBe(false);
    expect(needsSpacePreserve("")).toBe(false);
  });
});
