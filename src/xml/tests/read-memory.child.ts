// Run by read-memory.test.ts in a process of its own, with a heap limit and a collector to call:
// what it measures is the memory of that process. Prints one line of JSON, a `Report`.
import { getHeapStatistics } from "node:v8";
import { type ByteChunks, type XmlCursor, type XmlVisitor, readXml } from "../read.js";

/** A part made to exhaust a reader, and how reading it ended. */
export interface Flood {
  readonly name: string;
  readonly chunkBytes: number;
  /** `read`, or the code the part was refused with. */
  readonly outcome: string;
  /** How much of the part was asked for before it ended. */
  readonly bytesRead: number;
  readonly milliseconds: number;
  /** The most the heap held more than before the read, garbage collected, sampled as it was read. */
  readonly heapGrowth: number;
}

/** What a visitor kept of a large part, and what the heap holds for it once the part has ended. */
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
  /**
   * What the heap holds more, as a part ends, than before it was read, when one tag of the part
   * held as many attributes as a token may and was followed by 8 MiB of small elements.
   */
  readonly heldAfterAttributes: number;
}

const collect = (globalThis as { gc?: () => void }).gc ?? (() => undefined);

/** What the heap holds once garbage is gone. */
function held(): number {
  collect();
  collect();
  return process.memoryUsage().heapUsed;
}

const encoder = new TextEncoder();

/** One read in this many samples the heap. */
const SAMPLED_READS = 64;

interface Source extends ByteChunks {
  readonly bytesRead: number;
  readonly heapPeak: number;
}

/**
 * A part of `total` bytes that is never held whole: `head`, then `repeated` over and over, cut off
 * at `total`, in chunks of `chunkBytes`.
 */
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

/**
 * A part whose pieces `piece` makes one after the other, `total` bytes of them at the least, and
 * then `tail`.
 */
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

/** As large as a part may inflate to under the default limits. */
const FLOOD_BYTES = 134_217_728;
/** What a package gives of a stored part at a time, and the most an inflate gives of a deflated one. */
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
    made((index) => `<p:e${index}${"n".repeat(500_000)}>`, FLOOD_BYTES, SLICE),
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

/** A part of elements that each hold an attribute and a text, in every way those are written. */
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

/** Reads the part with a visitor that keeps something of one element in `EVERY`. */
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
  await readXml(source, { open: look("open"), close: look("close"), end: () => undefined });
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

const beforeAttributes = held();
let heldAfterAttributes = 0;
await readXml(
  made(
    (index) =>
      index === 0
        ? "<part><a"
        : index < 380_000
          ? ` n${index}=""`
          : index === 380_000
            ? "/>"
            : "<b><c>1</c></b>",
    12 * 1_048_576,
    SLICE,
    "</part>",
  ),
  {
    open() {},
    close: () => undefined,
    end() {
      heldAfterAttributes = held() - beforeAttributes;
    },
  },
);

const report: Report = {
  heapLimit: getHeapStatistics().heap_size_limit,
  floodBytes: FLOOD_BYTES,
  floods,
  partBytes: PART_BYTES,
  kept,
  heldAfterAttributes,
};
console.log(JSON.stringify(report));
