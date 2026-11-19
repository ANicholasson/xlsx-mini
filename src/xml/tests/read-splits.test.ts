import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { type Token, outcomeOf, tokensOf, utf8 } from "./reading.js";

// A part is generated as what it means, and each piece of it is then written in one of the ways
// XML allows for that meaning. The tokens expected of it come from the meaning alone.

type Place = "text" | "cdata" | '"' | "'";

/** One character of text or of an attribute value, and which of its ways to be written is taken. */
interface Character {
  readonly point: number;
  readonly way: number;
}

interface Attribute {
  readonly name: string;
  readonly quote: '"' | "'";
  readonly value: readonly Character[];
}

interface Element {
  readonly kind: "element";
  readonly name: string;
  readonly attributes: readonly Attribute[];
  readonly children: readonly Node[];
  /** Picks the white space inside the tags, and how an element without children is written. */
  readonly layout: number;
}

type Node =
  | Element
  | { readonly kind: "text" | "cdata"; readonly characters: readonly Character[] }
  | { readonly kind: "comment"; readonly body: string }
  | { readonly kind: "instruction"; readonly target: string; readonly body: string };

const ENTITIES: Readonly<Record<number, string>> = {
  38: "&amp;",
  60: "&lt;",
  62: "&gt;",
  34: "&quot;",
  39: "&apos;",
};

const LINE_ENDS = ["\n", "\r\n", "\r"];

/** Every way XML allows `point` to be written in `place` and still be read as `point`. */
function waysToWrite(point: number, place: Place): string[] {
  const character = String.fromCodePoint(point);
  if (place === "cdata") return point === 10 ? LINE_ENDS : [character];
  const hexadecimal = point.toString(16);
  const ways = [`&#${point};`, `&#x${hexadecimal};`, `&#x${hexadecimal.toUpperCase()};`];
  ways.push(`&#00${point};`, `&#x000${hexadecimal};`);
  const entity = ENTITIES[point];
  if (entity !== undefined) ways.push(entity);
  if (place === "text") {
    if (point === 10) ways.push(...LINE_ENDS);
    else if (point !== 13 && point !== 38 && point !== 60) ways.push(character);
    return ways;
  }
  // In an attribute value a tab and a line end written as themselves are read as a space.
  if (point === 32) ways.push(" ", "\t", "\n", "\r", "\r\n");
  else if (point > 32 && point !== 38 && point !== 60 && character !== place) ways.push(character);
  return ways;
}

function write(characters: readonly Character[], place: Place): string {
  const written = characters.map(({ point, way }) => {
    const ways = waysToWrite(point, place);
    return ways[way % ways.length]!;
  });
  // A carriage return before a line feed would make one line end of two characters.
  return written
    .map((piece, index) =>
      piece === "\r" && written[index + 1]?.startsWith("\n") ? "\r\n" : piece,
    )
    .join("");
}

function meaning(characters: readonly Character[]): string {
  return characters.map(({ point }) => String.fromCodePoint(point)).join("");
}

function isDeclaration(name: string): boolean {
  return name === "xmlns" || name.startsWith("xmlns:");
}

/** What an attribute is asked for by: its name without a prefix, a namespace declaration whole. */
function keyOf(name: string): string {
  return isDeclaration(name) ? name : name.slice(name.indexOf(":") + 1);
}

const SPACES = ["", " ", "\n", "\t ", "\r\n"];

function writeNode(node: Node): string {
  switch (node.kind) {
    case "text":
      // Character data may not hold the end of a CDATA section.
      return write(node.characters, "text").replaceAll("]]>", "]]&gt;");
    case "cdata":
      // Nor may a CDATA section: one that means it is written as two.
      return `<![CDATA[${write(node.characters, "cdata").replaceAll("]]>", "]]]]><![CDATA[>")}]]>`;
    case "comment":
      return `<!--${node.body}-->`;
    case "instruction":
      return `<?${node.target}${node.body}?>`;
    case "element": {
      const { layout } = node;
      const gap = (bits: number) => SPACES[(layout >> bits) % SPACES.length]!;
      const attributes = node.attributes
        .map(
          ({ name, quote, value }, index) =>
            `${gap(index) || " "}${name}${gap(index + 1)}=${gap(index + 2)}${quote}${write(value, quote)}${quote}`,
        )
        .join("");
      const open = `<${node.name}${attributes}${gap(3)}`;
      if (node.children.length === 0 && layout % 2 === 0) return `${open}/>`;
      // Two runs of text side by side are one run: they are written as one.
      const children: Node[] = [];
      for (const child of node.children) {
        const last = children.at(-1);
        if (child.kind === "text" && last?.kind === "text") {
          children[children.length - 1] = {
            kind: "text",
            characters: [...last.characters, ...child.characters],
          };
        } else {
          children.push(child);
        }
      }
      return `${open}>${children.map(writeNode).join("")}</${node.name}${gap(5)}>`;
    }
  }
}

function tokensOfModel(element: Element, depth: number, tokens: Token[]): Token[] {
  const name = element.name.slice(element.name.indexOf(":") + 1);
  const attributes: Record<string, string> = {};
  for (const attribute of element.attributes) {
    attributes[keyOf(attribute.name)] = meaning(attribute.value);
  }
  tokens.push(["open", name, depth, attributes]);
  let text = "";
  let leaf = true;
  for (const child of element.children) {
    if (child.kind === "element") {
      leaf = false;
      tokensOfModel(child, depth + 1, tokens);
    } else if (child.kind === "text" || child.kind === "cdata") {
      text += meaning(child.characters);
    }
  }
  tokens.push(["close", name, depth, leaf ? text : ""]);
  return tokens;
}

function keysOf(element: Element, keys = new Set<string>()): Set<string> {
  for (const attribute of element.attributes) keys.add(keyOf(attribute.name));
  for (const child of element.children) if (child.kind === "element") keysOf(child, keys);
  return keys;
}

const point = fc.oneof(
  { weight: 6, arbitrary: fc.integer({ min: 0x20, max: 0x7e }) },
  { weight: 3, arbitrary: fc.constantFrom(9, 10, 13, 32, 34, 38, 39, 60, 62, 93, 45, 63) },
  { weight: 2, arbitrary: fc.integer({ min: 0x7f, max: 0xd7ff }) },
  { weight: 1, arbitrary: fc.integer({ min: 0xe000, max: 0xfffd }) },
  { weight: 1, arbitrary: fc.integer({ min: 0x10000, max: 0x10ffff }) },
  { weight: 1, arbitrary: fc.constantFrom(0xe9, 0x20ac, 0x1f389, 0xfeff, 0x85, 0x2028, 0xd7ff) },
);
const characters = fc.array(fc.record({ point, way: fc.nat({ max: 60 }) }), { maxLength: 8 });

const elementName = fc.constantFrom(
  "a",
  "b",
  "c",
  "v",
  "t",
  "si",
  "row",
  "x:row",
  "r:id",
  "a-b.c_d",
);
const attributeName = fc.constantFrom(
  "r",
  "s",
  "t",
  "id",
  "r:id",
  "xml:space",
  "xmlns",
  "xmlns:r",
  "xmlns:x",
  "x14ac:dyDescent",
  "uniqueCount",
  "ref",
);

/** Text for a comment or an instruction, which XML passes over: without what would end it early. */
const passedOver = fc
  .array(point, { maxLength: 8 })
  .map((points) => String.fromCodePoint(...points).replaceAll("\r", " "));

const comment = passedOver.map((body): Node => ({
  kind: "comment",
  body: body.replaceAll("-", "- "),
}));
const instruction = fc
  .tuple(fc.constantFrom("p", "a-b", "xml-stylesheet", "xmlfoo", "xm"), passedOver)
  .map(([target, body]): Node => ({
    kind: "instruction",
    target,
    body: body === "" ? "" : ` ${body.replaceAll("?>", "? >")}`,
  }));
const text = characters.map((value): Node => ({ kind: "text", characters: value }));
const cdata = fc
  .array(fc.record({ point: point.filter((value) => value !== 13), way: fc.nat({ max: 60 }) }), {
    maxLength: 8,
  })
  .map((value): Node => ({ kind: "cdata", characters: value }));

const { element } = fc.letrec<{ element: Element; node: Node }>((tie) => ({
  element: fc.record({
    kind: fc.constant("element" as const),
    name: elementName,
    attributes: fc.uniqueArray(
      fc.record({
        name: attributeName,
        quote: fc.constantFrom('"' as const, "'" as const),
        value: characters,
      }),
      { maxLength: 11, selector: ({ name }) => keyOf(name) },
    ),
    children: fc.array(tie("node"), { maxLength: 4 }),
    layout: fc.nat({ max: 0xffff }),
  }),
  node: fc.oneof(
    { depthSize: "small" },
    { weight: 4, arbitrary: text },
    { weight: 1, arbitrary: cdata },
    { weight: 1, arbitrary: comment },
    { weight: 1, arbitrary: instruction },
    { weight: 4, arbitrary: tie("element") },
  ),
}));

const around = fc.array(fc.oneof(comment, instruction, fc.constantFrom(" ", "\n", "\r\n")), {
  maxLength: 2,
});

const declaration = fc.constantFrom(
  "",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
  `<?xml version='1.0' encoding='utf-8'?>`,
  `<?xml version = "1.0"\tstandalone = "no" ?>`,
);

interface Part {
  readonly root: Element;
  readonly bytes: Uint8Array;
}

const part: fc.Arbitrary<Part> = fc
  .record({
    mark: fc.constantFrom("", "\uFEFF"),
    declaration,
    before: around,
    root: element,
    after: around,
  })
  .map(({ mark, declaration: declared, before, root, after }) => {
    const misc = (nodes: readonly (Node | string)[]) =>
      nodes.map((node) => (typeof node === "string" ? node : writeNode(node))).join("");
    return {
      root,
      bytes: utf8(`${mark}${declared}${misc(before)}${writeNode(root)}${misc(after)}`),
    };
  });

/** Where to cut `length` bytes into chunks: anywhere, or after every byte. */
const cutsOf = (length: number): fc.Arbitrary<number[]> =>
  fc.oneof(
    { weight: 4, arbitrary: fc.array(fc.nat({ max: length }), { maxLength: 6 }) },
    { weight: 1, arbitrary: fc.constant(Array.from({ length }, (_, index) => index)) },
  );

function sorted(cuts: readonly number[]): number[] {
  return [...new Set(cuts)].sort((a, b) => a - b);
}

describe("a part cut into chunks", () => {
  it("gives the tokens its meaning has, wherever the bytes are cut", async () => {
    await fc.assert(
      fc.asyncProperty(
        part.chain((generated) => fc.tuple(fc.constant(generated), cutsOf(generated.bytes.length))),
        async ([{ root, bytes }, cuts]) => {
          const asked = [...keysOf(root)];
          const expected = tokensOfModel(root, 1, []);
          expect(await tokensOf(bytes, { asked })).toEqual(expected);
          expect(await tokensOf(bytes, { asked, cuts: sorted(cuts) })).toEqual(expected);
        },
      ),
      { numRuns: 400 },
    );
  });

  it("ends as it does in one chunk when bytes of it are changed: the same tokens, or refused", async () => {
    const edit = fc.record({
      at: fc.nat({ max: 4000 }),
      byte: fc.oneof(
        fc.constantFrom(
          ...utf8(`<>/&;"'= -?![]#x:\r\n`),
          0,
          0xff,
          0xc3,
          0xe2,
          0xf0,
          0xef,
          0xbb,
          0xbf,
        ),
        fc.nat({ max: 255 }),
      ),
      how: fc.constantFrom("put" as const, "drop" as const, "set" as const, "stop" as const),
    });
    await fc.assert(
      fc.asyncProperty(
        part.chain((generated) =>
          fc.tuple(
            fc.constant(generated),
            fc.array(edit, { minLength: 1, maxLength: 3 }),
            cutsOf(generated.bytes.length + 3),
          ),
        ),
        async ([{ root, bytes }, edits, cuts]) => {
          let changed = [...bytes];
          for (const { at, byte, how } of edits) {
            const index = at % (changed.length + 1);
            if (how === "put") changed.splice(index, 0, byte);
            else if (how === "drop") changed.splice(index, 1);
            else if (how === "set") changed[index] = byte;
            else changed = changed.slice(0, index);
          }
          const asked = [...keysOf(root)];
          const whole = await outcomeOf(new Uint8Array(changed), { asked });
          const cut = await outcomeOf(new Uint8Array(changed), { asked, cuts: sorted(cuts) });
          expect(cut).toEqual(whole);
        },
      ),
      { numRuns: 400 },
    );
  });

  const sample =
    `\uFEFF<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<!-- é - 🎉 --><?xml-model a="b"?>` +
    `<x:sst xmlns:x="urn:x" count="2"\tuniqueCount = '2'>\r\n` +
    `<x:si><x:t xml:space="preserve"> Tom &amp; Jerry &#x1F389;&#13;\r\n</x:t></x:si>` +
    `<x:si a="é&quot;€&#9;\t🎉"><x:t>«<![CDATA[<é>&amp;]]>»]]&gt;<!-- c -->€<?p ?>\u{10FFFF}</x:t><x:r/></x:si>` +
    `</x:sst >\n<!-- 🎉 -->`;

  const sampleTokens: Token[] = [
    ["open", "sst", 1, { "xmlns:x": "urn:x", count: "2", uniqueCount: "2" }],
    ["open", "si", 2, {}],
    ["open", "t", 3, { space: "preserve" }],
    ["close", "t", 3, " Tom & Jerry 🎉\r\n"],
    ["close", "si", 2, ""],
    ["open", "si", 2, { a: 'é"€\t 🎉' }],
    ["open", "t", 3, {}],
    ["close", "t", 3, "«<é>&amp;»]]>€\u{10FFFF}"],
    ["open", "r", 3, {}],
    ["close", "r", 3, ""],
    ["close", "si", 2, ""],
    ["close", "sst", 1, ""],
  ];
  const asked = ["xmlns:x", "count", "uniqueCount", "space", "a", "x"];

  it("gives the same tokens cut after any one byte, a multi-byte character included", async () => {
    const bytes = utf8(sample);
    expect(await tokensOf(bytes, { asked })).toEqual(sampleTokens);
    for (let cut = 0; cut <= bytes.length; cut++) {
      expect(await tokensOf(bytes, { asked, cuts: [cut] }), `cut after byte ${cut}`).toEqual(
        sampleTokens,
      );
    }
  });

  it("gives the same tokens a byte at a time, in threes, and with empty chunks between", async () => {
    const bytes = utf8(sample);
    const every = (step: number) =>
      Array.from({ length: Math.ceil(bytes.length / step) }, (_, index) => index * step);
    expect(await tokensOf(bytes, { asked, cuts: every(1) })).toEqual(sampleTokens);
    expect(await tokensOf(bytes, { asked, cuts: every(3) })).toEqual(sampleTokens);
    expect(await tokensOf(bytes, { asked, cuts: [0, 0, 1, 1, 1, 40, 40, bytes.length] })).toEqual(
      sampleTokens,
    );
  });

  it("refuses a second byte order mark wherever the first is cut", async () => {
    // Node 24 drops both marks when a chunk ends inside the first.
    const bytes = utf8("\uFEFF\uFEFF<a/>");
    for (let cut = 0; cut <= bytes.length; cut++) {
      expect(await outcomeOf(bytes, { cuts: [cut] }), `cut after byte ${cut}`).toBe("INVALID_XML");
    }
    expect(await outcomeOf(bytes, { cuts: [1, 2, 3, 4, 5, 6] })).toBe("INVALID_XML");
  });
});
