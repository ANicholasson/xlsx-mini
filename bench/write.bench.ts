import { describe, test } from "vite-plus/test";
import type * as Api from "../src/index.js";
import { manySheetsWorkbook, stringHeavyWorkbook, wideWorkbook } from "./workbooks.mjs";

// The built package, what users run, loaded outside the module runner; `pnpm bench` builds it
// first. The specifier is computed so a fresh checkout typechecks before any build.
const { writeXlsx } = (await import(
  new URL("../dist/index.js", import.meta.url).href
)) as typeof Api;

// CodSpeed's CPU simulation (and its memory profiler) executes every instruction under
// instrumentation, many times slower than native: one full-size write of the wide table takes
// minutes there. Under those modes the suites keep their shape at a twentieth of the rows, which
// still dwarfs the per-write overhead, and write STORE: DEFLATE runs in the runtime's native
// `CompressionStream` on other threads, which the simulation cannot measure, so those numbers would
// be waiting time rather than the writer's work. Locally and under walltime nothing changes.
const runnerMode =
  process.env.CODSPEED_ENV === undefined ? undefined : process.env.CODSPEED_RUNNER_MODE;
const instrumented =
  runnerMode === "simulation" || runnerMode === "instrumentation" || runnerMode === "memory";
const scale = instrumented ? 1 / 20 : 1;

// Built once, outside the timed functions: the benchmarks measure the writer, not the generator.
const wide = wideWorkbook(scale);
const manySheets = manySheetsWorkbook(scale);
const strings = stringHeavyWorkbook(scale);

/** Few, long iterations: one write takes long enough that more would only cost minutes. */
const runOptions = { iterations: 5, time: 0, warmupIterations: 1, warmupTime: 0 };

/** Room for the instrumented runs, which take far longer than the default timeout. */
const testOptions = { timeout: 600_000 };

/** The compression a single-variant suite writes with, and the benchmark name that says so. */
const compression = instrumented ? "none" : "deflate";
const compressionName = instrumented ? "store" : "deflate";

/** A row count as the suite names write it: `100 000`. */
function rows(count: number): string {
  return Math.round(count * scale)
    .toString()
    .replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

describe("writeXlsx", () => {
  test(`wide table: ${rows(100_000)} rows by 10 mixed columns`, testOptions, async ({ bench }) => {
    const store = bench("store", async () => {
      await writeXlsx(wide, { compression: "none" });
    });
    if (instrumented) {
      await store.run(runOptions);
      return;
    }
    await bench.compare(
      bench("deflate", async () => {
        await writeXlsx(wide);
      }),
      store,
      runOptions,
    );
  });

  test(
    `many sheets: 40 sheets of ${rows(2_500)} rows by 6 columns`,
    testOptions,
    async ({ bench }) => {
      await bench(compressionName, async () => {
        await writeXlsx(manySheets, { compression });
      }).run(runOptions);
    },
  );

  test(
    `string-heavy: ${rows(200_000)} rows by 5 text columns from 2 000 phrases`,
    testOptions,
    async ({ bench }) => {
      await bench(compressionName, async () => {
        await writeXlsx(strings, { compression });
      }).run(runOptions);
    },
  );
});
