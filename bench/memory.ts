// Peak memory per benchmark suite, each written once in a fresh Node process against dist/, and what
// reading holds: each suite read once from its bytes, the wide table read at 100 000 and 1 000 000
// rows from its bytes and from a `Blob` of its file, and a column of text kept from a large sheet.
// `pnpm bench:memory` prints the tables. `bun bench/memory.ts` and `deno run --allow-all
// bench/memory.ts` print the reads; a name as the argument runs one measurement.
import { execFile } from "node:child_process";
import { mkdtempSync, openAsBlob, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import type * as Api from "../src/index.ts";
import type * as Read from "../src/read.ts";
import {
  KEPT,
  type Kept,
  type KeptColumn,
  SAMPLED_ROWS,
  WIDE_LIMITS,
  bytesOf,
  held,
  keptColumn,
  keptWorkbook,
  peakOfReading,
} from "./read-memory.ts";
import { suites, wideStreamed } from "./workbooks.ts";

interface Measurement {
  readonly name: string;
  readonly outputBytes: number;
  readonly rssBeforeWrite: number;
  readonly peakRss: number;
  readonly milliseconds: number;
  /** Streamed only: the most heap still in use after a collection, sampled every 8 MiB of output. */
  readonly retainedHeap?: number;
}

/** The streamed measurements: the wide table at these row counts, pulled into a sink that discards the bytes. */
const STREAMED_ROWS = [100_000, 1_000_000] as const;

/** The rows of the file a read's process reads first. */
const SMALL_ROWS = 1_000;

const host = globalThis as {
  Bun?: { version: string };
  Deno?: { version: { deno: string } };
};

function fileBytes(path: string): Uint8Array {
  const { buffer, byteOffset, byteLength } = readFileSync(path);
  return new Uint8Array(buffer, byteOffset, byteLength);
}

const dist = (file: string) => new URL(`../dist/${file}`, import.meta.url).href;

/** One read in this process, from a file another process wrote: a suite, the wide table, or a kept column. */
interface ReadMeasurement {
  readonly name: string;
  readonly fileBytes: number;
  readonly rows: number;
  /** What was held before the read began: the package's bytes among it, when read from them. */
  readonly heldBefore: number;
  readonly peak: number;
  readonly milliseconds: number;
}

/**
 * One read, after a read of the small file `small` of the same kind, which compiles the reader's
 * code and leaves too little to sway what is measured when it is collected.
 */
async function measureRead(name: string, path: string, small: string): Promise<void> {
  const read = (await import(dist("read.js"))) as typeof Read;
  const [kind, detail] = name.split(":") as [string, string];
  if (kind === "kept") {
    await keptColumn(read, fileBytes(small), detail as Kept);
    const kept = await keptColumn(read, fileBytes(path), detail as Kept);
    console.log(JSON.stringify({ name, kept }));
    return;
  }
  const warm = await read.openXlsx(fileBytes(small));
  await peakOfReading([warm.rows(0)], SAMPLED_ROWS, held());
  const wide = kind === "wide";
  const input = detail.endsWith("blob") ? await openAsBlob(path) : fileBytes(path);
  const heldBefore = held();
  const started = performance.now();
  const book = await read.openXlsx(input, wide ? { limits: WIDE_LIMITS } : undefined);
  const { rows, peak } = await peakOfReading(
    book.sheets.map((_, sheet) => book.rows(sheet)),
    SAMPLED_ROWS,
    heldBefore,
  );
  const measurement: ReadMeasurement = {
    name,
    fileBytes: input instanceof Blob ? input.size : input.length,
    rows,
    heldBefore,
    peak,
    milliseconds: performance.now() - started,
  };
  console.log(JSON.stringify(measurement));
}

/** Streams the wide table with `rows` rows, produced as they are pulled, and reports peak RSS. */
async function measureStreamed(rows: number): Promise<void> {
  const { streamXlsx } = (await import(dist("index.js"))) as typeof Api;
  (globalThis as { gc?: () => void }).gc?.();
  const rssBeforeWrite = process.memoryUsage().rss;
  const started = performance.now();
  let outputBytes = 0;
  let retainedHeap = 0;
  let nextSample = 0;
  await streamXlsx(wideStreamed(rows)).pipeTo(
    new WritableStream({
      write(chunk) {
        outputBytes += chunk.length;
        if (outputBytes < nextSample) return;
        // The resident set grows with V8's appetite, not with what is held: collect, then look.
        (globalThis as { gc?: () => void }).gc?.();
        retainedHeap = Math.max(retainedHeap, process.memoryUsage().heapUsed);
        nextSample += 8 * 1_048_576;
      },
    }),
  );
  const milliseconds = performance.now() - started;
  const peakRss = process.resourceUsage().maxRSS * 1024;
  const measurement: Measurement = {
    name: `streamed ${rows / 1000}k`,
    outputBytes,
    rssBeforeWrite,
    peakRss,
    milliseconds,
    retainedHeap,
  };
  console.log(JSON.stringify(measurement));
}

/** Builds the workbook, collects garbage, then writes it and reports the process's peak RSS. */
async function measure(name: string): Promise<void> {
  const { writeXlsx } = (await import(dist("index.js"))) as typeof Api;
  const entry = suites.find(([suiteName]) => suiteName === name);
  if (entry === undefined) throw new Error(`unknown suite ${name}`);
  const workbook = entry[1]();
  (globalThis as { gc?: () => void }).gc?.();
  const rssBeforeWrite = process.memoryUsage().rss;
  const started = performance.now();
  const bytes = await writeXlsx(workbook);
  const milliseconds = performance.now() - started;
  const peakRss = process.resourceUsage().maxRSS * 1024;
  const measurement: Measurement = {
    name,
    outputBytes: bytes.byteLength,
    rssBeforeWrite,
    peakRss,
    milliseconds,
  };
  console.log(JSON.stringify(measurement));
}

const mb = (bytes: number) => (bytes / 1_048_576).toFixed(1).padStart(8);

/** This runtime again, with a collector to call, for one measurement in a process of its own. */
function engine(): readonly string[] {
  if (host.Bun !== undefined) return [];
  if (host.Deno !== undefined) return ["run", "--allow-all", "--no-lock", "--v8-flags=--expose-gc"];
  return ["--expose-gc"];
}

async function report(): Promise<void> {
  const run = promisify(execFile);
  const self = fileURLToPath(import.meta.url);
  const measured = async <Result>(...args: string[]): Promise<Result> => {
    const { stdout } = await run(process.execPath, [...engine(), self, ...args], {
      maxBuffer: 1 << 20,
    });
    return JSON.parse(stdout) as Result;
  };
  const runtime =
    host.Bun !== undefined
      ? `Bun ${host.Bun.version}`
      : host.Deno !== undefined
        ? `Deno ${host.Deno.version.deno}`
        : `Node ${process.versions.node}`;
  console.log(runtime);

  if (host.Bun === undefined && host.Deno === undefined) {
    const rows: Measurement[] = [];
    const names = [...suites.map(([name]) => name), ...STREAMED_ROWS.map((n) => `stream:${n}`)];
    for (const name of names) rows.push(await measured<Measurement>(name));
    console.log(
      "suite           output MiB  write ms  RSS before write  peak RSS  growth MiB  growth / output  retained heap MiB",
    );
    for (const row of rows) {
      const growth = row.peakRss - row.rssBeforeWrite;
      const retained = row.retainedHeap === undefined ? "—".padStart(8) : mb(row.retainedHeap);
      console.log(
        `${row.name.padEnd(14)}${mb(row.outputBytes)}  ${String(Math.round(row.milliseconds)).padStart(8)}  ${mb(row.rssBeforeWrite).padStart(16)}  ${mb(row.peakRss)}  ${mb(growth).padStart(10)}  ${(growth / row.outputBytes).toFixed(1).padStart(15)}x  ${retained.padStart(17)}`,
      );
    }
  }

  // Every file is written here, so a read's process holds nothing a write left behind.
  const api = (await import(dist("index.js"))) as typeof Api;
  const folder = mkdtempSync(join(tmpdir(), "xlsx-mini-memory-"));
  process.on("exit", () => rmSync(folder, { recursive: true, force: true }));
  const written = async (name: string, bytes: Uint8Array | Promise<Uint8Array>) => {
    const path = join(folder, `${name}.xlsx`);
    writeFileSync(path, await bytes);
    return path;
  };
  const small = await written("wide small", bytesOf(api.streamXlsx(wideStreamed(SMALL_ROWS))));
  const reads: ReadMeasurement[] = [];
  for (const [name, make] of suites) {
    const path = await written(name, api.writeXlsx(make()));
    reads.push(await measured<ReadMeasurement>(`read:${name}`, path, small));
  }
  for (const rows of STREAMED_ROWS) {
    const path = await written(`wide ${rows}`, bytesOf(api.streamXlsx(wideStreamed(rows))));
    for (const source of ["bytes", "blob"]) {
      reads.push(await measured<ReadMeasurement>(`wide:${rows / 1000}k ${source}`, path, small));
    }
  }
  console.log(
    "read                    file MiB    rows  read ms  held before MiB  peak above it MiB",
  );
  for (const row of reads) {
    console.log(
      `${row.name.padEnd(22)}${mb(row.fileBytes)}  ${String(row.rows).padStart(7)}  ${String(Math.round(row.milliseconds)).padStart(7)}  ${mb(row.heldBefore).padStart(15)}  ${mb(row.peak).padStart(17)}`,
    );
  }

  console.log(
    "kept column       strings  characters  wrong  part MiB  retained MiB  bytes per character  workbook collected",
  );
  for (const kept of KEPT) {
    const keptSmall = await written(
      `kept ${kept} small`,
      bytesOf(api.streamXlsx(...keptWorkbook(kept, SMALL_ROWS))),
    );
    const path = await written(
      `kept ${kept}`,
      bytesOf(api.streamXlsx(...keptWorkbook(kept, 100_000))),
    );
    const { kept: column } = await measured<{ kept: KeptColumn }>(`kept:${kept}`, path, keptSmall);
    console.log(
      `${kept.padEnd(16)}${String(column.strings).padStart(8)}  ${String(column.characters).padStart(10)}  ${String(column.wrong).padStart(5)}  ${mb(column.partBytes)}  ${mb(column.retained).padStart(12)}  ${(column.retained / column.characters).toFixed(1).padStart(19)}  ${String(column.collected).padStart(18)}`,
    );
  }
}

if ((import.meta as { main?: boolean }).main === true) {
  const name = process.argv[2];
  if (name === undefined) await report();
  else if (name.startsWith("stream:")) await measureStreamed(Number(name.slice("stream:".length)));
  else if (/^(read|wide|kept):/.test(name)) {
    await measureRead(name, process.argv[3]!, process.argv[4]!);
  } else await measure(name);
}
