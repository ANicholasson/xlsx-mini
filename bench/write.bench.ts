import { describe, test } from "vite-plus/test";
import type * as Api from "../src/index.js";
import { manySheetsWorkbook, stringHeavyWorkbook, wideWorkbook } from "./workbooks.mjs";

// The built package, what users run, loaded outside the module runner; `pnpm bench` builds it
// first. The specifier is computed so a fresh checkout typechecks before any build.
const { writeXlsx } = (await import(
  new URL("../dist/index.js", import.meta.url).href
)) as typeof Api;

// Built once, outside the timed functions: the benchmarks measure the writer, not the generator.
const wide = wideWorkbook();
const manySheets = manySheetsWorkbook();
const strings = stringHeavyWorkbook();

/** Few, long iterations: one write takes long enough that more would only cost minutes. */
const runOptions = { iterations: 5, time: 0, warmupIterations: 1, warmupTime: 0 };

describe("writeXlsx", () => {
  test("wide table: 100 000 rows by 10 mixed columns", async ({ bench }) => {
    await bench.compare(
      bench("deflate", async () => {
        await writeXlsx(wide);
      }),
      bench("store", async () => {
        await writeXlsx(wide, { compression: "none" });
      }),
      runOptions,
    );
  });

  test("many sheets: 40 sheets of 2 500 rows by 6 columns", async ({ bench }) => {
    await bench("deflate", async () => {
      await writeXlsx(manySheets);
    }).run(runOptions);
  });

  test("string-heavy: 200 000 rows by 5 text columns from 2 000 phrases", async ({ bench }) => {
    await bench("deflate", async () => {
      await writeXlsx(strings);
    }).run(runOptions);
  });
});
