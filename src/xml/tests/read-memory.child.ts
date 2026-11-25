// Run by read-memory.test.ts in a process of its own, with a heap limit and a collector to call.
// Prints one line of JSON, a `Report`.
import { getHeapStatistics } from "node:v8";
import { type ByteChunks, type XmlCursor, type XmlVisitor, readXml } from "../read.js";

export interface Flood {
  readonly name: string;
  readonly chunkBytes: number;
  readonly outcome: string;
  readonly bytesRead: number;
  readonly milliseconds: number;
  readonly heapGrowth: number;
}

interface Keeping {
  readonly what: string;
  readonly strings: number;
  readonly characters: number;
  readonly heapGrowth: number;
}

export interface Report {
  readonly heapLimit: number;
  readonly floodBytes: number;
  readonly floods: readonly Flood[];
  readonly partBytes: number;
  readonly kept: readonly Keeping[];
  readonly heldAfterAttributes: number;
  readonly heldAtNextTag: number;
  readonly heldAtNextTagPrefixed: number;
  readonly nested: number;
  readonly heldInScope: number;
}

const collect = (globalThis as { gc?: () => void }).gc ?? (() => undefined);

function held(): number {
  collect();
  collect();
  return process.memoryUsage().heapUsed;
}

const encoder = new TextEncoder();

const SAMPLED_READS = 64;

interface Source extends ByteChunks {
  readonly bytesRead: number;
  readonly heapPeak: number;
}

function repeating(head: string, repeated: string, total: number, chunkBytes: number): Source {
  const start = encoder.encode(head);
  const unit = encoder.encode(repeated);
  const body = new Uint8Array(chunkBytes + unit.length);
  for (let at = 0; at < body.length; at++) body[at] = unit[at % unit.length]!;
  let bytesRead = 0;
  let reads = 0;
  let heapPeak = 0;
  return {
    get bytesRead() {
      return bytesRead;
    },
    get heapPeak() {
      return heapPeak;
    },
    read() {
      if (reads++ % SAMPLED_READS === 0) heapPeak = Math.max(heapPeak, held());
      if (bytesRead >= total) return Promise.resolve(undefined);
      const length = Math.min(chunkBytes, total - bytesRead);
      let chunk: Uint8Array;
      if (bytesRead === 0) {
        chunk = new Uint8Array(length);
        chunk.set(start);
        chunk.set(body.subarray(0, length - start.length), start.length);
      } else {
        const phase = (bytesRead - start.length) % unit.length;
        chunk = body.subarray(phase, phase + length);
      }
      bytesRead += length;
      return Promise.resolve(chunk);
    },
    close() {},
  };
}

function made(
  piece: (index: number) => string,
  total: number,
  chunkBytes: number,
  tail = "",
): Source {
  let bytesRead = 0;
  let reads = 0;
  let index = 0;
  let ended = false;
  let heapPeak = 0;
  return {
    get bytesRead() {
      return bytesRead;
    },
    get heapPeak() {
      return heapPeak;
    },
    read() {
      if (reads++ % SAMPLED_READS === 0) heapPeak = Math.max(heapPeak, held());
      if (ended) return Promise.resolve(undefined);
      let text = "";
      while (text.length < chunkBytes && bytesRead + text.length < total) text += piece(index++);
      if (bytesRead + text.length >= total) {
        ended = true;
        text += tail;
      }
      const chunk = encoder.encode(text);
      bytesRead += chunk.length;
      return Promise.resolve(chunk);
    },
    close() {},
  };
}

const counting: XmlVisitor<never, string> = {
  namespaces: new Map(),
  open() {},
  close: () => undefined,
  end: () => "read",
};

async function flood(name: string, chunkBytes: number, source: Source): Promise<Flood> {
  const before = held();
  const started = performance.now();
  const outcome = await readXml(source, counting).catch(
    (error: unknown) => (error as { code: string }).code,
  );
  const milliseconds = performance.now() - started;
  return {
    name,
    chunkBytes,
    outcome,
    bytesRead: source.bytesRead,
    milliseconds,
    heapGrowth: Math.max(source.heapPeak, held()) - before,
  };
}

const FLOOD_BYTES = 134_217_728;
const SLICE = 16_384;
const LARGEST_CHUNK = 16_777_216;

const floods: Flood[] = [];
const repeated: [name: string, head: string, unit: string][] = [
  ["open tags", "", "<a>"],
  ["one run of text", "<a>", "x"],
  ["one attribute value", '<a b="', "x"],
  ["one name", "<", "a"],
  ["white space in one tag", "<a", " "],
  ["white space in one close tag", "<a></a", " "],
  ["one comment", "<a><!--", "x"],
  ["one CDATA section", "<a><![CDATA[", "x"],
  ["one instruction", "<a><?p ", "x"],
  ["one reference", "<a>&", "a"],
  ["the text of one element, in runs between comments", "<a>", "xxxxxxxxxxxxxxxx<!---->"],
  ["white space after the root", "<a/>", " "],
];
for (const [name, head, unit] of repeated) {
  floods.push(await flood(name, SLICE, repeating(head, unit, FLOOD_BYTES, SLICE)));
}
for (const [name, head, unit] of repeated.slice(0, 3)) {
  floods.push(await flood(name, LARGEST_CHUNK, repeating(head, unit, FLOOD_BYTES, LARGEST_CHUNK)));
}
floods.push(
  await flood(
    "attributes of one tag",
    SLICE,
    made((index) => (index === 0 ? "<a" : ` n${index}=""`), FLOOD_BYTES, SLICE),
  ),
  await flood(
    "the long names of 256 nested elements",
    SLICE,
    made(
      (index) => `<p:e${index}${"n".repeat(500_000)}${index === 0 ? ' xmlns:p="urn:p"' : ""}>`,
      FLOOD_BYTES,
      SLICE,
    ),
  ),
  await flood(
    "the long namespace names of nested declarations",
    SLICE,
    made((index) => `<e xmlns:p="urn:${"n".repeat(500_000)}${index}">`, FLOOD_BYTES, SLICE),
  ),
  await flood(
    "a declaration on every element, each out of scope when the next comes",
    SLICE,
    made(
      (index) => (index === 0 ? "<r>" : `<p:e xmlns:p="urn:p${index % 1000}" p:a="${index}"/>`),
      FLOOD_BYTES,
      SLICE,
      "</r>",
    ),
  ),
  await flood(
    "the text of one element, in runs between long comments",
    SLICE,
    made(
      (index) => (index === 0 ? "<a>" : `<!--${"c".repeat(16_000)}-->${"x".repeat(16)}`),
      FLOOD_BYTES,
      SLICE,
      "</a>",
    ),
  ),
);

const PART_BYTES = 67_108_864;
/**
 * One element in this many is kept: more often than a chunk ends, so that every chunk is kept
 * from, and a prime, so that every way of writing an element is.
 */
const EVERY = 127;

const TEXTS = ["", " &amp; more", " <!-- c -->more", " <![CDATA[more]]>"];

function element(index: number): string {
  const number = String(index).padStart(9, "0");
  const value = index % 2 === 0 ? "" : " &amp; more";
  return `<record-${number} k="attribute ${number}${value}">text ${number}${TEXTS[index % 4]!}</record-${number}>`;
}

const keepers: [what: string, on: "open" | "close", keep: (xml: XmlCursor) => string][] = [
  ["the text of an element", "close", (xml) => xml.text()],
  ["the value of an attribute", "open", (xml) => xml.attribute("k") ?? ""],
  ["the name of an element", "open", (xml) => xml.name],
  [
    "a slice of the buffer, which a visitor must not keep",
    "open",
    (xml) => xml.buffer.slice(xml.valueStart("k"), xml.valueEnd),
  ],
];

async function keeping(
  what: string,
  on: "open" | "close",
  keep: (xml: XmlCursor) => string,
): Promise<Keeping> {
  const before = held();
  const strings: string[] = [];
  let seen = 0;
  const look = (at: "open" | "close") => (xml: XmlCursor) => {
    if (at === on && xml.depth === 2 && seen++ % EVERY === 0) strings.push(keep(xml));
    return undefined;
  };
  const source = made(
    (index) => (index === 0 ? "<part>" : element(index)),
    PART_BYTES,
    SLICE,
    "</part>",
  );
  await readXml(source, {
    namespaces: new Map(),
    open: look("open"),
    close: look("close"),
    end: () => undefined,
  });
  const heapGrowth = held() - before;
  return {
    what,
    strings: strings.length,
    characters: strings.reduce((sum, text) => sum + text.length, 0),
    heapGrowth,
  };
}

const kept: Keeping[] = [];
for (const [what, on, keep] of keepers) kept.push(await keeping(what, on, keep));

/**
 * What is held once a tag of as many attributes as a token may hold is followed by others: at the
 * end of the part, or when the next tag opens, which still holds the text fed after the long tag.
 */
async function heldAfter(count: number, prefix: string, atNextTag: boolean): Promise<number> {
  const before = held();
  let after = Number.NaN;
  await readXml(
    made(
      (index) =>
        index === 0
          ? '<part xmlns:p="urn:p"><a'
          : index < count
            ? ` ${prefix}n${index}=""`
            : index === count
              ? "/>"
              : "<b><c>1</c></b>",
      12 * 1_048_576,
      SLICE,
      "</part>",
    ),
    {
      namespaces: new Map(),
      open(xml) {
        if (atNextTag && xml.name === "b" && Number.isNaN(after)) after = held() - before;
      },
      close: () => undefined,
      end() {
        if (!atNextTag) after = held() - before;
      },
    },
  );
  return after;
}

const heldAfterAttributes = await heldAfter(380_000, "", false);
const heldAtNextTag = await heldAfter(300_000, "", true);
const heldAtNextTagPrefixed = await heldAfter(300_000, "p:", true);

const NESTED = 200;

function nestedDeclarations(): ByteChunks {
  const opened = Array.from(
    { length: NESTED },
    (_, index) =>
      `<e xmlns:prefix${"p".repeat(16)}${index}="urn:${"u".repeat(64)}:${index}">${" ".repeat(SLICE)}`,
  );
  const bytes = encoder.encode(`${opened.join("")}<deepest/>${"</e>".repeat(NESTED)}`);
  let at = 0;
  return {
    read() {
      const chunk = at < bytes.length ? bytes.subarray(at, at + SLICE) : undefined;
      at += SLICE;
      return Promise.resolve(chunk);
    },
    close() {},
  };
}

const declarations = nestedDeclarations();
const beforeScope = held();
let heldInScope = Number.NaN;
await readXml(declarations, {
  namespaces: new Map(),
  open(xml) {
    if (xml.depth === NESTED + 1) heldInScope = held() - beforeScope;
  },
  close: () => undefined,
  end: () => undefined,
});

const report: Report = {
  heapLimit: getHeapStatistics().heap_size_limit,
  floodBytes: FLOOD_BYTES,
  floods,
  partBytes: PART_BYTES,
  kept,
  heldAfterAttributes,
  heldAtNextTag,
  heldAtNextTagPrefixed,
  nested: NESTED,
  heldInScope,
};
console.log(JSON.stringify(report));
