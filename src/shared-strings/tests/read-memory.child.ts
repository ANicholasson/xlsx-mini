// Run by read-memory.test.ts in a process of its own, with a small heap. Prints one line of JSON, a
// `Report`.
import { resolveLimits } from "../../errors/read.js";
import { type ByteChunks, readXml } from "../../xml/read.js";
import { sharedStringsVisitor } from "../read.js";

export interface Table {
  readonly name: string;
  readonly partBytes: number;
  readonly entries: number;
  readonly characters: number;
  /** Characters of the table other than the one the runs hold at their place. */
  readonly wrong: number;
}

export interface Report {
  readonly tables: readonly Table[];
}

const collect = (globalThis as { gc?: () => void }).gc ?? (() => undefined);
const encoder = new TextEncoder();

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const PART_BYTES = 33_554_432;
const CHUNK_BYTES = 16_384;

const HEAD = encoder.encode(`<sst xmlns="${MAIN}">`);
const OPEN = encoder.encode("<si>");
const CLOSE = encoder.encode("</si>");
const TAIL = encoder.encode("</sst>");

function* chunks(items: number, unit: string, runs: number): Generator<Uint8Array, undefined> {
  const perChunk = Math.floor(CHUNK_BYTES / unit.length);
  const body = encoder.encode(unit.repeat(perChunk));
  yield HEAD;
  for (let index = 0; index < items; index++) {
    yield OPEN;
    for (let left = runs; left > 0; left -= perChunk) {
      yield body.subarray(0, Math.min(left, perChunk) * unit.length);
    }
    yield CLOSE;
  }
  yield TAIL;
  return undefined;
}

/** A table of `items` items of runs of `run`, as many as the default `sharedStringBytes` holds. */
function part(items: number, run: string): { readonly bytes: ByteChunks; readonly size: number } {
  const unit = `<r><t>${run}</t></r>`;
  const runs = Math.floor((PART_BYTES / items - 128) / unit.length);
  const read = chunks(items, unit, runs);
  return {
    size: HEAD.length + items * (OPEN.length + runs * unit.length + CLOSE.length) + TAIL.length,
    bytes: { read: () => Promise.resolve(read.next().value), close() {} },
  };
}

const tables: Table[] = [];
for (const [name, items, run] of [
  ["one item of runs of 13 characters", 1, "abcdefghijklm"],
  ["one item of runs of one character", 1, "a"],
  ["two thousand items of runs of 16 characters", 2_000, "abcdefghijklmnop"],
] as const) {
  collect();
  const { bytes, size } = part(items, run);
  const table = await readXml(bytes, sharedStringsVisitor(resolveLimits(undefined).book.strings));
  let characters = 0;
  let wrong = 0;
  for (const text of table) {
    for (let at = 0; at < text.length; at++) {
      if (text.charCodeAt(at) !== run.charCodeAt(at % run.length)) wrong++;
    }
    characters += text.length;
  }
  tables.push({ name, partBytes: size, entries: table.length, characters, wrong });
}

const report: Report = { tables };
console.log(JSON.stringify(report));
