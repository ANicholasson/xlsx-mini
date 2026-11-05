// Peak memory per benchmark suite, each written once in a fresh Node process against dist/.
// `pnpm bench:memory` prints the table; a suite name as the argument runs one measurement.
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { suites } from "./workbooks.mjs";

const suite = process.argv[2];
if (suite === undefined) {
  await report();
} else {
  await measure(suite);
}

/** Builds the workbook, collects garbage, then writes it and reports the process's peak RSS. */
async function measure(name) {
  const { writeXlsx } = await import("../dist/index.js");
  const build = suites[name];
  if (build === undefined) throw new Error(`unknown suite ${name}`);
  const workbook = build();
  globalThis.gc?.();
  const rssBeforeWrite = process.memoryUsage().rss;
  const started = performance.now();
  const bytes = await writeXlsx(workbook);
  const milliseconds = performance.now() - started;
  const peakRss = process.resourceUsage().maxRSS * 1024;
  console.log(
    JSON.stringify({ name, outputBytes: bytes.byteLength, rssBeforeWrite, peakRss, milliseconds }),
  );
}

async function report() {
  const run = promisify(execFile);
  const self = fileURLToPath(import.meta.url);
  const rows = [];
  for (const name of Object.keys(suites)) {
    const { stdout } = await run(process.execPath, ["--expose-gc", self, name], {
      maxBuffer: 1 << 20,
    });
    rows.push(JSON.parse(stdout));
  }
  const mb = (bytes) => (bytes / 1_048_576).toFixed(1).padStart(8);
  console.log("suite         output MiB  write ms  RSS before write  peak RSS  growth / output");
  for (const row of rows) {
    const growth = row.peakRss - row.rssBeforeWrite;
    console.log(
      `${row.name.padEnd(12)}${mb(row.outputBytes)}  ${String(Math.round(row.milliseconds)).padStart(8)}  ${mb(row.rssBeforeWrite).padStart(16)}  ${mb(row.peakRss)}  ${(growth / row.outputBytes).toFixed(1).padStart(15)}x`,
    );
  }
}
