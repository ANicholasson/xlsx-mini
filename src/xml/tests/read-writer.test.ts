import { readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { keptByXml } from "../../../test/helpers/text.js";
import { unzip } from "../../../test/helpers/unzip.js";
import { escapeXml } from "../index.js";
import { type Token, outcomeOf, tokensOf } from "./reading.js";

const fixtures = join(
  dirname(fileURLToPath(import.meta.url)),
  "..",
  "..",
  "..",
  "test",
  "fixtures",
);

const REFERENCE = /&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g;
const NAMED: Readonly<Record<string, string>> = {
  lt: "<",
  gt: ">",
  amp: "&",
  quot: '"',
  apos: "'",
};

function withoutReferences(text: string): string {
  return text.replace(REFERENCE, (_, body: string) => {
    if (!body.startsWith("#")) return NAMED[body]!;
    return String.fromCodePoint(
      Number(body.startsWith("#x") ? `0${body.slice(1)}` : body.slice(1)),
    );
  });
}

const MARKUP = /<\?xml [^?]*\?>|<(\/?)([\w.:-]+)((?:\s+[\w.:-]+="[^"<]*")*)\s*(\/?)>/g;

/**
 * The tokens of a part the writer wrote, found with patterns: the writer writes every attribute in
 * double quotes and escapes every `<` and `>` it did not put there itself. A reader that shares
 * nothing with the tokenizer.
 */
function tokensByPattern(xml: string): { tokens: Token[]; asked: string[] } {
  const tokens: Token[] = [];
  const asked = new Set<string>();
  const open: { text: string; leaf: boolean }[] = [];
  let read = 0;
  for (const match of xml.matchAll(MARKUP)) {
    const between = xml.slice(read, match.index);
    expect(between).not.toContain("<");
    read = match.index + match[0].length;
    const [, slash, name, attributes, empty] = match;
    const parent = open.at(-1);
    if (parent === undefined) expect(between.trim()).toBe("");
    else parent.text += between;
    if (name === undefined || attributes === undefined) continue;
    const local = name.slice(name.indexOf(":") + 1);
    if (slash === "/") {
      const closed = open.pop()!;
      const text = withoutReferences(closed.text.replace(/\r\n?/g, "\n"));
      tokens.push(["close", local, open.length + 1, closed.leaf ? text : ""]);
      continue;
    }
    if (parent !== undefined) parent.leaf = false;
    const found: Record<string, string> = {};
    for (const [, written, value] of attributes.matchAll(/([\w.:-]+)="([^"]*)"/g)) {
      const declaration = written === "xmlns" || written!.startsWith("xmlns:");
      const key = declaration ? written! : written!.slice(written!.indexOf(":") + 1);
      asked.add(key);
      found[key] = withoutReferences(value!.replace(/\r\n?|[\n\t]/g, " "));
    }
    tokens.push(["open", local, open.length + 1, found]);
    if (empty === "/") tokens.push(["close", local, open.length + 1, ""]);
    else open.push({ text: "", leaf: true });
  }
  expect(xml.slice(read).trim()).toBe("");
  expect(open).toEqual([]);
  return { tokens, asked: [...asked] };
}

describe("what the writer writes", () => {
  it("reads back from every part of every golden fixture as a reader by patterns reads it, whole and in chunks", async () => {
    let parts = 0;
    const files = readdirSync(fixtures).filter((name) => name.endsWith(".xlsx"));
    for (const file of files) {
      for (const entry of unzip(readFileSync(join(fixtures, file)))) {
        if (!/\.(xml|rels)$/.test(entry.name)) continue;
        parts++;
        const where = `${file}: ${entry.name}`;
        const { tokens, asked } = tokensByPattern(new TextDecoder().decode(entry.data));
        expect(tokens.length, where).toBeGreaterThan(1);
        expect(await tokensOf(entry.data, { asked }), where).toEqual(tokens);
        const cuts = Array.from({ length: 12 }, (_, index) =>
          Math.floor(((index + 1) * entry.data.length) / 13),
        );
        expect(await tokensOf(entry.data, { asked, cuts }), where).toEqual(tokens);
      }
    }
    expect(files.length).toBeGreaterThan(20);
    expect(parts).toBeGreaterThanOrEqual(files.length * 6);
  });

  it("reads back as text and as an attribute value from what escapeXml makes of any text", async () => {
    const anyCodeUnits = fc
      .array(fc.integer({ min: 0, max: 0xffff }), { maxLength: 40 })
      .map((units) => String.fromCharCode(...units));
    const special = fc
      .array(fc.constantFrom("\r", "\n", "\t", " ", "&", "<", ">", '"', "'", "]", "\u0000", "🎉"), {
        maxLength: 12,
      })
      .map((pieces) => pieces.join(""));
    await fc.assert(
      fc.asyncProperty(
        fc.oneof(fc.string({ unit: "binary" }), anyCodeUnits, special),
        async (text) => {
          const escaped = escapeXml(text);
          const kept = keptByXml(text);
          expect(await tokensOf(`<t a="${escaped}">${escaped}</t>`, { asked: ["a"] })).toEqual([
            ["open", "t", 1, { a: kept.replace(/\r\n?|[\n\t]/g, " ") }],
            ["close", "t", 1, kept.replace(/\r\n?/g, "\n")],
          ]);
        },
      ),
      { numRuns: 300 },
    );
  });

  const keeps = (point: number): boolean => escapeXml(String.fromCodePoint(point)) !== "";

  it("is read character for character: every character the writer keeps, written as itself", async () => {
    let all = "";
    for (let point = 0; point <= 0x10ffff; point++) {
      if (keeps(point) && point !== 38 && point !== 60) all += String.fromCodePoint(point);
    }
    expect(all.length).toBeGreaterThan(2_000_000);
    const tokens = await tokensOf(`<t a="${all.replace('"', "'")}">${all}</t>`, { asked: ["a"] });
    expect(tokens[1]![3] === all.replace("\r", "\n")).toBe(true);
    const value = all.replace('"', "'").replace(/[\t\n\r]/g, " ");
    expect((tokens[0]![3] as Record<string, string>)["a"] === value).toBe(true);
  });

  it("is read character for character: a reference to every character the writer keeps", async () => {
    const elements: string[] = [];
    const texts: string[] = [];
    for (let from = 0; from <= 0x10ffff; from += 0x20000) {
      let written = "";
      let text = "";
      for (let point = from; point < from + 0x20000 && point <= 0x10ffff; point++) {
        if (!keeps(point)) continue;
        written += point % 2 === 0 ? `&#${point};` : `&#x${point.toString(16)};`;
        text += String.fromCodePoint(point);
      }
      elements.push(`<t>${written}</t>`);
      texts.push(text);
    }
    const tokens = await tokensOf(`<r>${elements.join("")}</r>`);
    const read = tokens.filter(([kind, name]) => kind === "close" && name === "t");
    expect(read).toHaveLength(texts.length);
    expect(read.every(([, , , text], index) => text === texts[index])).toBe(true);
  });

  it("is all that is read: a character the writer drops is refused, as itself and as a reference", async () => {
    const dropped: number[] = [];
    for (let point = 0; point <= 0xffff; point++) if (!keeps(point)) dropped.push(point);
    // The control characters but three, the surrogates, and the last two code points of the plane.
    expect(dropped).toHaveLength(29 + 2048 + 2);
    for (const point of [...dropped, 0x110000, 0x7fffffff]) {
      const reference = `&#x${point.toString(16)};`;
      expect(await outcomeOf(`<t>${reference}</t>`), reference).toBe("INVALID_XML");
      expect(await outcomeOf(`<t a="&#${point};"/>`), reference).toBe("INVALID_XML");
      // A surrogate alone has no bytes in UTF-8 to be written as itself with.
      if (point >= 0xd800 && point <= 0xdfff) continue;
      if (point > 0xffff) continue;
      const character = String.fromCharCode(point);
      expect(await outcomeOf(`<t>${character}</t>`), reference).toBe("INVALID_XML");
      expect(await outcomeOf(`<t a="${character}"/>`), reference).toBe("INVALID_XML");
    }
  });
});
