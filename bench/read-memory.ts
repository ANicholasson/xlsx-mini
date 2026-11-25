// What reading holds, measured the same way by `bench/memory.ts` on Node, Bun and Deno and by the
// read memory test under a heap cap: the most a read holds above what was held before it, and what
// a column of text kept from a large sheet holds once the rest is collected.
import type * as Api from "../src/index.ts";
import type * as Read from "../src/read.ts";
import { unzip } from "../test/helpers/unzip.ts";

const host = globalThis as { gc?: () => void; Bun?: { gc: (force: boolean) => void } };

function collect(): void {
  if (host.gc !== undefined) {
    host.gc();
    host.gc();
  } else {
    host.Bun?.gc(true);
  }
}

/**
 * Collects garbage, then gives what the engine holds: V8's heap and the array buffers outside it,
 * or JavaScriptCore's heap, which counts its array buffers. Deno reports no array buffers.
 */
export function held(): number {
  collect();
  const { heapUsed, arrayBuffers } = process.memoryUsage();
  return host.Bun === undefined ? heapUsed + arrayBuffers : heapUsed;
}

/**
 * Collects garbage, then gives the engine's heap alone, where text is: V8 counts an array buffer it
 * has collected until it frees the buffer's memory, in its own time, which moved a column's figure
 * by 4.5 MiB on Node 24.
 */
function heapHeld(): number {
  collect();
  return process.memoryUsage().heapUsed;
}

/** What a read held at most: above what was held before it, after a collection every `every` rows. */
export interface ReadPeak {
  readonly rows: number;
  readonly peak: number;
}

/**
 * Reads every row of `sheets`, one sheet after another, collecting every `every` rows. Measure a
 * read alone in its process, after a read of a small file that compiles the reader's code: what an
 * earlier read leaves can still be held when this one starts, and be collected while it runs.
 */
export async function peakOfReading(
  sheets: readonly AsyncIterable<unknown>[],
  every: number,
  before: number,
): Promise<ReadPeak> {
  let rows = 0;
  let peak = 0;
  for (const sheet of sheets) {
    for await (const _ of sheet) {
      if (++rows % every === 0) peak = Math.max(peak, held() - before);
    }
  }
  return { rows, peak: Math.max(peak, held() - before) };
}

/** Where the text of a kept column comes from: each way a cell holds text, and objects by a header. */
export const KEPT = ["shared strings", "inline strings", "formula text", "objects"] as const;
export type Kept = (typeof KEPT)[number];

/** Longer than the 13 characters from which V8 cuts a string as a view of the one it is cut from. */
const keptText = (index: number) => `kept ${String(index).padStart(12, "0")} z`;
/** Text no one keeps: 25 times the kept text, so a part held whole outweighs the column. */
const filler = (index: number) => `${String(index).padStart(10, "0")} ${"filler text ".repeat(40)}`;

/** A sheet under a header row whose first column is kept and whose second is not, and how to stream it. */
export function keptWorkbook(
  kept: Kept,
  rows: number,
): readonly [Api.StreamedWorkbook, Api.StreamOptions] {
  function* generated(): Generator<Api.Cell[]> {
    yield ["Kept", "Filler"];
    for (let index = 0; index < rows; index++) {
      const text = keptText(index);
      yield [kept === "formula text" ? { formula: `"${text}"`, value: text } : text, filler(index)];
    }
  }
  return [
    { sheets: [{ name: "Kept", rows: generated() }] },
    { strings: kept === "shared strings" ? "shared" : "inline" },
  ];
}

/** Room for the shared strings of the kept sheet, which hold the filler text too. */
export const KEPT_LIMITS = { sharedStringBytes: 64 * 1_048_576 } as const;

/**
 * One kept column: how many strings and characters, how many differ from the text written, what
 * letting go of the column frees, and whether the workbook was collected, without which its shared
 * strings would stay whatever the column holds.
 */
export interface KeptColumn {
  readonly strings: number;
  readonly characters: number;
  readonly wrong: number;
  readonly retained: number;
  readonly collected: boolean;
  /** What the part the text was read from inflates to. */
  readonly partBytes: number;
}

async function keptTexts(
  read: Pick<typeof Read, "openXlsx" | "objectsFromRows">,
  file: Uint8Array,
  kept: Kept,
): Promise<{ readonly texts: string[]; readonly workbook: WeakRef<object> }> {
  const book = await read.openXlsx(file, { limits: KEPT_LIMITS });
  const texts: string[] = [];
  if (kept === "objects") {
    const columns = [{ key: "kept", header: "Kept" }] as const;
    for await (const { values } of read.objectsFromRows(book.rows(0), columns)) {
      texts.push(String(values.kept));
    }
  } else {
    for await (const { number, cells } of book.rows(0)) {
      if (number > 1) texts.push(String(cells[0]));
    }
  }
  return { texts, workbook: new WeakRef(book) };
}

/** A turn of the event loop, after which nothing of the read is on the stack. */
const turn = () => new Promise((resolve) => setTimeout(resolve, 0));

/**
 * Reads the first column of a kept sheet and keeps its text and nothing else, then lets go of it:
 * what the column held is what a collection frees then. Measured so, nothing else that is collected
 * between the two looks counts, and a column that held its chunks frees them too.
 */
export async function keptColumn(
  read: Pick<typeof Read, "openXlsx" | "objectsFromRows">,
  file: Uint8Array,
  kept: Kept,
): Promise<KeptColumn> {
  const { texts, workbook } = await keptTexts(read, file, kept);
  await turn();
  const holding = heapHeld();
  const collected = workbook.deref() === undefined;
  const strings = texts.length;
  let characters = 0;
  let wrong = 0;
  for (const [index, text] of texts.entries()) {
    characters += text.length;
    if (text !== keptText(index)) wrong++;
  }
  texts.length = 0;
  await turn();
  const retained = holding - heapHeld();
  const part = kept === "shared strings" ? "xl/sharedStrings.xml" : "xl/worksheets/sheet1.xml";
  const partBytes = unzip(file).find(({ name }) => name === part)!.uncompressedSize;
  return { strings, characters, wrong, retained, collected, partBytes };
}

/** What a million rows of the wide table need beyond the defaults: 493 MiB inflated, 10 million positions. */
export const WIDE_LIMITS = { inflatedBytes: 1024 * 1_048_576, cellPositions: 20_000_000 } as const;

/** How often a read collects and samples what is held. */
export const SAMPLED_ROWS = 10_000;

export async function bytesOf(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
