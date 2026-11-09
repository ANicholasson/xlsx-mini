// Peak memory per benchmark suite, each written once in a fresh Node process against dist/.
// `pnpm bench:memory` prints the table; a suite name as the argument runs one measurement.
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import type * as Api from "../src/index.ts";
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

const suite = process.argv[2];
if (suite === undefined) {
  await report();
} else if (suite.startsWith("stream:")) {
  await measureStreamed(Number(suite.slice("stream:".length)));
} else {
  await measure(suite);
}

/** Streams the wide table with `rows` rows, produced as they are pulled, and reports peak RSS. */
async function measureStreamed(rows: number): Promise<void> {
  const { streamXlsx } = (await import(
    new URL("../dist/index.js", import.meta.url).href
  )) as typeof Api;
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
  const { writeXlsx } = (await import(
    new URL("../dist/index.js", import.meta.url).href
  )) as typeof Api;
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

async function report(): Promise<void> {
  const run = promisify(execFile);
  const self = fileURLToPath(import.meta.url);
  const rows: Measurement[] = [];
  const names = [...suites.map(([name]) => name), ...STREAMED_ROWS.map((n) => `stream:${n}`)];
  for (const name of names) {
    const { stdout } = await run(process.execPath, ["--expose-gc", self, name], {
      maxBuffer: 1 << 20,
    });
    rows.push(JSON.parse(stdout) as Measurement);
  }
  const mb = (bytes: number) => (bytes / 1_048_576).toFixed(1).padStart(8);
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
