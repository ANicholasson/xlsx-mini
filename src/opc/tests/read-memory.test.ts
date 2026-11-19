import { execFile } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { rolldown } from "rolldown";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import type { Opening, Report } from "./read-memory.child.ts";

const here = dirname(fileURLToPath(import.meta.url));

const MEBIBYTE = 1_048_576;

/** The old space the child's heap is held to, in mebibytes. */
const HEAP_MEBIBYTES = 512;
/** Milliseconds within which a package is opened or refused. */
const OPENING_MILLISECONDS = 1_000;
/** The most the heap may hold more after a refusal than before it. */
const REFUSAL_HEAP_BYTES = 4 * MEBIBYTE;
/** The most the heap may hold for an opened package of 90 000 parts: 279 bytes a part. */
const OPENED_HEAP_BYTES = 24 * MEBIBYTE;
/**
 * What is read of a file to find its directory: an end record with the longest comment, and the
 * zip64 end record and its locator in front of it.
 */
const END_OF_FILE_BYTES = 22 + 65_535 + 56 + 20;
/** The most the process may hold more while it reads a part than before it. */
const READING_BYTES = 8 * MEBIBYTE;

describe("the package reader in a process whose heap is held to 512 MiB", () => {
  let report: Report;

  const openingsOf = (entries: number): readonly Opening[] =>
    report.openings.filter((opening) => opening.entries === entries);

  beforeAll(async () => {
    const directory = mkdtempSync(join(tmpdir(), "xlsx-mini-read-"));
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
        [`--max-old-space-size=${HEAP_MEBIBYTES}`, "--expose-gc", script, directory],
        { encoding: "utf8" },
      );
      report = JSON.parse(stdout) as Report;
    } finally {
      // Removed here, once the child has ended and holds its file no longer.
      rmSync(directory, { recursive: true, force: true });
    }
  }, 120_000);

  it("runs with that limit: its whole heap, the young space included, is under a gibibyte", () => {
    expect(report.heapLimit).toBeLessThan(1024 * MEBIBYTE);
  });

  it("refuses a package of four million entries by the size of its directory, from bytes and from a Blob, without reading the directory", () => {
    const openings = openingsOf(4_000_000);
    expect(openings.map(({ from }) => from)).toEqual(["bytes", "a Blob"]);
    for (const { from, outcome, milliseconds, heapGrowth, bytesRead } of openings) {
      expect(outcome, from).toEqual({ code: "LIMIT_EXCEEDED", limit: "inflatedBytes" });
      expect(milliseconds, from).toBeLessThan(OPENING_MILLISECONDS);
      expect(heapGrowth, from).toBeLessThan(REFUSAL_HEAP_BYTES);
      expect(bytesRead, from).toBe(from === "a Blob" ? END_OF_FILE_BYTES : 0);
    }
  });

  it("refuses a package of 110 000 entries by their number", () => {
    const openings = openingsOf(110_000);
    expect(openings.map(({ from }) => from)).toEqual(["bytes", "a Blob"]);
    for (const { from, outcome, milliseconds, heapGrowth } of openings) {
      expect(outcome, from).toEqual({ code: "LIMIT_EXCEEDED", limit: "parts" });
      expect(milliseconds, from).toBeLessThan(OPENING_MILLISECONDS);
      expect(heapGrowth, from).toBeLessThan(REFUSAL_HEAP_BYTES);
    }
  });

  it("opens a package of 90 000 entries and reads its last part", () => {
    const openings = openingsOf(90_000);
    expect(openings.map(({ from }) => from)).toEqual(["bytes", "a Blob"]);
    for (const { from, outcome, milliseconds, heapGrowth } of openings) {
      expect(outcome, from).toBe("opened");
      expect(milliseconds, from).toBeLessThan(OPENING_MILLISECONDS);
      expect(heapGrowth, from).toBeLessThan(OPENED_HEAP_BYTES);
    }
  });

  it("reads a part of 128 MiB in stored blocks from a Blob over a file and holds less than 8 MiB of it at a time", () => {
    const { fileBytes, contentBytes, retainedGrowth } = report.readings.stored;
    expect(contentBytes).toBe(128 * MEBIBYTE);
    expect(fileBytes).toBeGreaterThan(contentBytes);
    expect(retainedGrowth).toBeLessThan(READING_BYTES);
  });

  it("reads a deflated part of 128 MiB from a Blob over a file and holds less than 8 MiB of it at a time", () => {
    const { fileBytes, contentBytes, retainedGrowth } = report.readings.deflated;
    expect(contentBytes).toBe(128 * MEBIBYTE);
    expect(fileBytes).toBeLessThan(contentBytes / 4);
    expect(retainedGrowth).toBeLessThan(READING_BYTES);
  });
});
