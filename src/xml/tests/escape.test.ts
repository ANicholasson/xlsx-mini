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
});

describe("needsSpacePreserve", () => {
  it("is true only for leading or trailing whitespace", () => {
    expect(needsSpacePreserve(" padded")).toBe(true);
    expect(needsSpacePreserve("padded\t")).toBe(true);
    expect(needsSpacePreserve("in the middle")).toBe(false);
    expect(needsSpacePreserve("")).toBe(false);
  });
});
