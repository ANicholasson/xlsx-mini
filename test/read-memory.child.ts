// Run by read-memory.test.ts in a process of its own, with a heap limit and a collector to call:
// `flat <file> <bytes|blob> <small file>` reads the wide table and prints a `FlatRead`, and
// `kept <file> <kept> <small file>` keeps a column and prints a `KeptRead`. Each reads the small
// file first, so the reader's code is compiled before what is measured.
import { openAsBlob, readFileSync } from "node:fs";
import { getHeapStatistics } from "node:v8";
import {
  type Kept,
  type KeptColumn,
  type ReadPeak,
  SAMPLED_ROWS,
  WIDE_LIMITS,
  baseline,
  keptColumn,
  peakOfReading,
} from "../bench/read-memory.ts";
import { objectsFromRows, openXlsx } from "../src/read.js";

export interface FlatRead extends ReadPeak {
  readonly heapLimit: number;
}

export interface KeptRead extends KeptColumn {
  readonly heapLimit: number;
}

function fileBytes(path: string): Uint8Array {
  const { buffer, byteOffset, byteLength } = readFileSync(path);
  return new Uint8Array(buffer, byteOffset, byteLength);
}

/** Every row of the one sheet of `input`, with the limits a million rows of the wide table need. */
async function readOnce(input: Uint8Array | Blob): Promise<ReadPeak> {
  const book = await openXlsx(input, { limits: WIDE_LIMITS });
  return peakOfReading([book.rows(0)], SAMPLED_ROWS, await baseline());
}

const [mode, path, what, small] = process.argv.slice(2) as [string, string, string, string];
const heapLimit = getHeapStatistics().heap_size_limit;
if (mode === "flat") {
  await readOnce(fileBytes(small));
  const input = what === "blob" ? await openAsBlob(path) : fileBytes(path);
  console.log(JSON.stringify({ heapLimit, ...(await readOnce(input)) } satisfies FlatRead));
} else {
  const read = { openXlsx, objectsFromRows };
  await keptColumn(read, fileBytes(small), what as Kept);
  const column = await keptColumn(read, fileBytes(path), what as Kept);
  console.log(JSON.stringify({ heapLimit, ...column } satisfies KeptRead));
}
