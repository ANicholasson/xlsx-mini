import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { reportOf } from "./helpers/child.ts";
import type { Report } from "./read-memory.child.ts";

const here = dirname(fileURLToPath(import.meta.url));

const MEBIBYTE = 1_048_576;
const HEAP_MEBIBYTES = 512;

describe("rows in a process whose heap is held to 512 MiB", () => {
  let report: Report;

  beforeAll(async () => {
    report = (await reportOf(join(here, "read-memory.child.ts"), HEAP_MEBIBYTES)) as Report;
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
