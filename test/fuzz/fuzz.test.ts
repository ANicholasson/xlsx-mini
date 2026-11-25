import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import {
  CHECKED_RUN,
  type FuzzFailure,
  type FuzzReport,
  ISSUE_MARKER,
  fuzz,
  issueOf,
  logLine,
} from "./fuzz.ts";

const here = dirname(fileURLToPath(import.meta.url));

describe(`${CHECKED_RUN.count} mutations of the corpus from seed ${CHECKED_RUN.seed}`, () => {
  let report: FuzzReport;
  beforeAll(async () => {
    report = await fuzz(CHECKED_RUN);
  }, 120_000);

  it("end in a refusal with a code or a clean read, and none reads a sheet that holds a value without rows", () => {
    expect(report.failures.map(logLine)).toEqual([]);
    expect(report.ended.size).toBe(CHECKED_RUN.count);
  });

  it("reach every depth, each with clean reads and refusals of more than one code", () => {
    for (const [depth, { mutations, read, readOther, refused }] of Object.entries(report.tallies)) {
      expect(mutations, depth).toBeGreaterThan(CHECKED_RUN.count / 5);
      expect(read + readOther, depth).toBeGreaterThan(mutations / 50);
      expect(Object.keys(refused).length, depth).toBeGreaterThan(2);
    }
    expect(report.tallies.text.readOther).toBeGreaterThan(0);
    expect(report.tallies.text.excused).toBeGreaterThan(0);
  });

  it.each([0, 1_234, CHECKED_RUN.count - 1])(
    "read mutation %i again, alone, from its seed and its index, to the same end",
    async (index) => {
      const again = await fuzz({ seed: CHECKED_RUN.seed, first: index, count: 1 });
      expect(again.ended).toEqual(new Map([[index, report.ended.get(index)]]));
    },
    60_000,
  );

  it("writes an issue the scheduled job finds again, with the log line of each failure", () => {
    const failure: FuzzFailure = {
      seed: 7,
      index: 0,
      workbook: "excel/minimal.xlsx",
      depth: "text",
      mutation: "xl/workbook.xml: delete <sheets> at 40",
      failure: "not an XlsxReadError: TypeError",
    };
    const failures = Array.from({ length: 60 }, (_, index) => ({ ...failure, index }));
    const text = issueOf({ ...report, seed: 7, failures }, "https://example.invalid/run");
    expect(text.startsWith(ISSUE_MARKER)).toBe(true);
    expect(text).toContain("https://example.invalid/run");
    expect(text).toContain(logLine(failures[49]!));
    expect(text).toContain("node test/fuzz/fuzz.ts --seed 7 --mutation 49");
    expect(text).not.toContain("--mutation 50");
    expect(text).toContain("and 10 more");
    const workflow = readFileSync(
      join(here, "..", "..", ".github", "workflows", "fuzz.yml"),
      "utf8",
    );
    expect(workflow).toContain(`startswith("${ISSUE_MARKER}")`);
  });
});
