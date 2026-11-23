import { execFile } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { rolldown } from "rolldown";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import type { Report } from "./read-memory.child.ts";

const here = dirname(fileURLToPath(import.meta.url));

const MEBIBYTE = 1_048_576;
const HEAP_MEBIBYTES = 512;

describe("rows in a process whose heap is held to 512 MiB", () => {
  let report: Report;

  beforeAll(async () => {
    const directory = mkdtempSync(join(tmpdir(), "xlsx-mini-rows-"));
    try {
      // The child is TypeScript, which Node 20 does not run: bundled, it is one file it does.
      const build = await rolldown({
        input: join(here, "read-memory.child.ts"),
        platform: "node",
        logLevel: "silent",
      });
      const { output } = await build.generate({ format: "esm" });
      await build.close();
      const script = join(directory, "child.mjs");
      writeFileSync(script, output[0].code);
      const { stdout } = await promisify(execFile)(
        process.execPath,
        [`--max-old-space-size=${HEAP_MEBIBYTES}`, script],
        { encoding: "utf8" },
      );
      report = JSON.parse(stdout) as Report;
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  }, 120_000);

  it("refuses three sheets of one row of 119 MiB each, read at once and inside every other default limit, with rowCharacters", () => {
    expect(report.heapLimit).toBeLessThan(1024 * MEBIBYTE);
    expect(report.readings.map(({ row }) => row)).toEqual(["cells", "runs"]);
    for (const { row, sheetBytes, fileBytes, outcomes, milliseconds } of report.readings) {
      expect(sheetBytes, row).toBeGreaterThan(118 * MEBIBYTE);
      expect(sheetBytes, row).toBeLessThanOrEqual(128 * MEBIBYTE);
      expect(fileBytes, row).toBeLessThan(MEBIBYTE);
      expect(outcomes, row).toEqual([
        "LIMIT_EXCEEDED rowCharacters Sheet1",
        "LIMIT_EXCEEDED rowCharacters Sheet2",
        "LIMIT_EXCEEDED rowCharacters Sheet3",
      ]);
      expect(milliseconds, row).toBeLessThan(5_000);
    }
  });
});
