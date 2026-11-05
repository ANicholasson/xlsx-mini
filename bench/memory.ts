// Peak memory per benchmark suite, each written once in a fresh Node process against dist/.
// `pnpm bench:memory` prints the table; a suite name as the argument runs one measurement.
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import type * as Api from "../src/index.ts";
import { suites } from "./workbooks.ts";

interface Measurement {
  readonly name: string;
  readonly outputBytes: number;
  readonly rssBeforeWrite: number;
  readonly peakRss: number;
  readonly milliseconds: number;
}

const suite = process.argv[2];
if (suite === undefined) {
  await report();
} else {
  await measure(suite);
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
  for (const [name] of suites) {
    const { stdout } = await run(process.execPath, ["--expose-gc", self, name], {
      maxBuffer: 1 << 20,
    });
    rows.push(JSON.parse(stdout) as Measurement);
  }
  const mb = (bytes: number) => (bytes / 1_048_576).toFixed(1).padStart(8);
  console.log("suite         output MiB  write ms  RSS before write  peak RSS  growth / output");
  for (const row of rows) {
    const growth = row.peakRss - row.rssBeforeWrite;
    console.log(
      `${row.name.padEnd(12)}${mb(row.outputBytes)}  ${String(Math.round(row.milliseconds)).padStart(8)}  ${mb(row.rssBeforeWrite).padStart(16)}  ${mb(row.peakRss)}  ${(growth / row.outputBytes).toFixed(1).padStart(15)}x`,
    );
  }
}
