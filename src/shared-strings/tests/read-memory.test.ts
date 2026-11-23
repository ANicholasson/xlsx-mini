import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { reportOf } from "../../../test/helpers/child.ts";
import type { Report } from "./read-memory.child.ts";

const here = dirname(fileURLToPath(import.meta.url));

const HEAP_MEBIBYTES = 64;
const DEFAULT_TABLE_BYTES = 33_554_432;

describe("the shared-string table in a process whose heap is held to 64 MiB", () => {
  let report: Report;

  beforeAll(async () => {
    report = (await reportOf(join(here, "read-memory.child.ts"), HEAP_MEBIBYTES)) as Report;
  }, 120_000);

  it("holds an item of many runs as its characters, for a part of the largest size the default limits allow", () => {
    expect(report.tables.map(({ name, entries, wrong }) => ({ name, entries, wrong }))).toEqual([
      { name: "one item of runs of 13 characters", entries: 1, wrong: 0 },
      { name: "one item of runs of one character", entries: 1, wrong: 0 },
      { name: "two thousand items of runs of 16 characters", entries: 2_000, wrong: 0 },
    ]);
    for (const { name, partBytes, characters } of report.tables) {
      expect(partBytes, name).toBeLessThanOrEqual(DEFAULT_TABLE_BYTES);
      expect(partBytes, name).toBeGreaterThan(DEFAULT_TABLE_BYTES - 1_000_000);
      expect(characters, name).toBeGreaterThan(0);
    }
  });
});
