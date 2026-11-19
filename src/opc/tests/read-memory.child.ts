// Run by read-memory.test.ts in a process of its own, with a heap limit and a collector to call:
// what it measures is the memory of that process. Takes a folder to write one file in, and prints
// one line of JSON, a `Report`.
import { closeSync, openAsBlob, openSync, writeSync } from "node:fs";
import { join } from "node:path";
import { getHeapStatistics } from "node:v8";
import { resolveLimits } from "../../errors/read.js";
import { type StreamedPart, selectStreamCompression, streamPackage } from "../index.js";
import { type PackageReader, openPackage, partName } from "../read.js";
import { noise } from "./reading.js";

/** How one package fared: opened, or the refusal it ended in. */
export interface Opening {
  readonly entries: number;
  readonly from: "bytes" | "a Blob";
  readonly outcome: "opened" | { readonly code: string; readonly limit: string | undefined };
  readonly milliseconds: number;
  /**
   * What the heap holds more after the call than before it, garbage collected both times: for a
   * package that opened, the opened package.
   */
  readonly heapGrowth: number;
  /** For a `Blob`: every byte read from it. */
  readonly bytesRead: number;
}

/** A large package read from a file through a `Blob`. */
interface Reading {
  readonly fileBytes: number;
  readonly contentBytes: number;
  /** The most the process held more than before the read, sampled every 16 MiB of content. */
  readonly retainedGrowth: number;
  readonly milliseconds: number;
}

export interface Report {
  readonly heapLimit: number;
  readonly openings: readonly Opening[];
  /** A part of noise in stored blocks, and a part of rows that the platform's deflate shrinks. */
  readonly readings: { readonly stored: Reading; readonly deflated: Reading };
}

const collect = (globalThis as { gc?: () => void }).gc ?? (() => undefined);

/**
 * What the process holds once garbage is gone: its heap, and the buffers outside the heap. A buffer
 * is freed a turn of the event loop after the collection that found it, so there are two.
 */
async function held(): Promise<{ readonly heap: number; readonly buffers: number }> {
  collect();
  await new Promise((resolve) => setImmediate(resolve));
  collect();
  const { heapUsed, external } = process.memoryUsage();
  return { heap: heapUsed, buffers: external };
}

/** A `Blob` that counts the bytes read from it. */
class CountedBlob extends Blob {
  bytesRead = 0;

  override slice(start = 0, end = this.size): Blob {
    this.bytesRead += Math.min(end, this.size) - start;
    return super.slice(start, end);
  }
}

const NAME_BYTES = 7;
const LOCAL_BYTES = 30 + NAME_BYTES;
const CENTRAL_BYTES = 46 + NAME_BYTES;

const nameOf = (entry: number) => entry.toString(36).padStart(NAME_BYTES, "0");

/**
 * A package of `entries` empty stored parts, each with a name and a local header of its own, and
 * with the zip64 end record that more than 65 535 entries need.
 */
function flood(entries: number): Uint8Array<ArrayBuffer> {
  const directoryStart = entries * LOCAL_BYTES;
  const directorySize = entries * CENTRAL_BYTES;
  const bytes = new Uint8Array(directoryStart + directorySize + 56 + 20 + 22);
  const view = new DataView(bytes.buffer);
  for (let entry = 0; entry < entries; entry++) {
    const name = nameOf(entry);
    const local = entry * LOCAL_BYTES;
    const central = directoryStart + entry * CENTRAL_BYTES;
    view.setUint32(local, 0x04034b50, true);
    view.setUint16(local + 4, 20, true);
    view.setUint16(local + 26, NAME_BYTES, true);
    view.setUint32(central, 0x02014b50, true);
    view.setUint16(central + 4, 20, true);
    view.setUint16(central + 6, 20, true);
    view.setUint16(central + 28, NAME_BYTES, true);
    view.setUint32(central + 42, local, true);
    for (let at = 0; at < NAME_BYTES; at++) {
      const code = name.charCodeAt(at);
      bytes[local + 30 + at] = code;
      bytes[central + 46 + at] = code;
    }
  }
  const record = directoryStart + directorySize;
  view.setUint32(record, 0x06064b50, true);
  view.setBigUint64(record + 4, 44n, true);
  view.setUint16(record + 12, 45, true);
  view.setUint16(record + 14, 45, true);
  view.setBigUint64(record + 24, BigInt(entries), true);
  view.setBigUint64(record + 32, BigInt(entries), true);
  view.setBigUint64(record + 40, BigInt(directorySize), true);
  view.setBigUint64(record + 48, BigInt(directoryStart), true);
  const locator = record + 56;
  view.setUint32(locator, 0x07064b50, true);
  view.setBigUint64(locator + 8, BigInt(record), true);
  view.setUint32(locator + 16, 1, true);
  const end = locator + 20;
  view.setUint32(end, 0x06054b50, true);
  view.setUint16(end + 8, 0xffff, true);
  view.setUint16(end + 10, 0xffff, true);
  view.setUint32(end + 12, 0xffffffff, true);
  view.setUint32(end + 16, 0xffffffff, true);
  return bytes;
}

async function opening(entries: number, file: Uint8Array | CountedBlob): Promise<Opening> {
  const limits = resolveLimits(undefined);
  const before = await held();
  const started = performance.now();
  let pkg: PackageReader | undefined;
  let outcome: Opening["outcome"] = "opened";
  try {
    pkg = await openPackage(file, limits.book.parts, limits.book.structure);
  } catch (error) {
    const { code, limit } = error as { code: string; limit: string | undefined };
    outcome = { code, limit };
  }
  const milliseconds = performance.now() - started;
  const heapGrowth = (await held()).heap - before.heap;
  const last = pkg?.open(partName(nameOf(entries - 1), "worksheet"), limits.rows().inflate);
  if (pkg !== undefined && (last === undefined || (await last.read()) !== undefined)) {
    throw new Error("the last part of an opened package is there and empty");
  }
  return {
    entries,
    from: file instanceof Blob ? "a Blob" : "bytes",
    outcome,
    milliseconds,
    heapGrowth,
    bytesRead: file instanceof Blob ? file.bytesRead : 0,
  };
}

/** 64 KiB of rows as a worksheet holds them, which a deflate shrinks about five times. */
function rows(): Uint8Array {
  let text = "";
  for (let row = 1; text.length < 65_536; row++) {
    text += `<row r="${row}"><c r="A${row}"><v>${row * 7}</v></c><c r="B${row}" t="s"><v>${row % 97}</v></c></row>`;
  }
  return new TextEncoder().encode(text.slice(0, 65_536));
}

/** A part of `bytes` bytes, `chunk` over and over, made as it is pulled. */
function body(bytes: number, chunk: Uint8Array): ReadableStream<Uint8Array> {
  let made = 0;
  return new ReadableStream({
    pull(controller) {
      if (made >= bytes) return controller.close();
      made += chunk.length;
      controller.enqueue(chunk);
    },
  });
}

/**
 * Writes a package of one part of `contentBytes` bytes to a file, as `streamPackage` writes it in
 * stored blocks or through the platform's deflate, and reads the part back through a `Blob` over
 * that file.
 */
async function reading(
  directory: string,
  contentBytes: number,
  compression: "none" | "deflate",
): Promise<Reading> {
  const path = join(directory, `large-${compression}.xlsx`);
  const sheet = "xl/worksheets/sheet1.xml";
  const file = openSync(path, "w");
  const chunk = compression === "none" ? noise(65_536) : rows();
  async function* parts(): AsyncGenerator<StreamedPart> {
    yield { path: sheet, body: body(contentBytes, chunk) };
  }
  for await (const written of streamPackage(parts(), selectStreamCompression(compression))) {
    writeSync(file, written);
  }
  closeSync(file);

  const blob = await openAsBlob(path);
  const limits = resolveLimits(undefined);
  const before = await held();
  const started = performance.now();
  const pkg = await openPackage(blob, limits.book.parts, limits.book.structure);
  const part = pkg.open(partName(sheet, "worksheet"), limits.rows().inflate);
  let read = 0;
  let sampled = 0;
  let retainedGrowth = 0;
  for (let chunk = await part?.read(); chunk !== undefined; chunk = await part?.read()) {
    read += chunk.length;
    if (read - sampled >= 16 * 1_048_576) {
      sampled = read;
      const now = await held();
      retainedGrowth = Math.max(
        retainedGrowth,
        now.heap + now.buffers - before.heap - before.buffers,
      );
    }
  }
  if (read !== contentBytes) throw new Error(`read ${read} bytes of ${contentBytes}`);
  return {
    fileBytes: blob.size,
    contentBytes,
    retainedGrowth,
    milliseconds: performance.now() - started,
  };
}

const openings: Opening[] = [];
for (const entries of [4_000_000, 110_000, 90_000]) {
  const bytes = flood(entries);
  openings.push(await opening(entries, bytes), await opening(entries, new CountedBlob([bytes])));
}

const [directory] = process.argv.slice(2);
if (directory === undefined) throw new Error("no folder to write in");
// As large as a part may be under the default limits.
const report: Report = {
  heapLimit: getHeapStatistics().heap_size_limit,
  openings,
  readings: {
    stored: await reading(directory, 134_217_728, "none"),
    deflated: await reading(directory, 134_217_728, "deflate"),
  },
};
console.log(JSON.stringify(report));
