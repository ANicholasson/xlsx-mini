import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import { KEPT, type Kept, bytesOf, keptWorkbook } from "../bench/read-memory.ts";
import { wideStreamed } from "../bench/workbooks.ts";
import { streamXlsx } from "../src/index.js";
import { reportOf } from "./helpers/child.ts";
import type { FlatRead, KeptRead } from "./read-memory.child.ts";

const child = join(dirname(fileURLToPath(import.meta.url)), "read-memory.child.ts");

const MEBIBYTE = 1_048_576;
/**
 * What a heap may hold besides its old space, which the cap limits: the young space, 96 MiB on
 * Node 26 and 192 MiB on Node 24 on the machine that set this. Without the cap the heap is gibibytes.
 */
const YOUNG_SPACE_BYTES = 256 * MEBIBYTE;
/** The rows of the file each child reads first, so that its own read meets compiled code. */
const SMALL_ROWS = 1_000;

const FLAT_ROWS = [100_000, 1_000_000] as const;
const SOURCES = ["bytes", "blob"] as const;
/** A read of the wide table holds 0.6 to 1.4 MiB above what was held before it, on Node 24 and 26. */
const FLAT_HEAP_MEBIBYTES = 64;
/**
 * A read holds a row and the text of its part around it, so a million rows hold what 100 000 hold.
 * What is held between rows varies by a few hundred kibibytes, and ten times the samples meet more
 * of it: the ratio of the two peaks was 1.21 to 1.38 on Node 24 and 26. A reader that keeps the
 * number of every row it read reaches 5.8 to 6.7.
 */
const FLAT_RATIO = 2;

const KEPT_ROWS = 100_000;
/** Room for the shared strings of the kept sheet, which the workbook holds: a part of 54 MiB. */
const KEPT_HEAP_MEBIBYTES = 256;
/**
 * Twice what a kept text of 19 characters takes on Node 24 and 26, 4.3 bytes per character: a copy
 * of it and a view of the copy. A column that held its chunks would free 32 to 38 bytes per kept
 * character with them, as much as its part, which holds 24 characters or more for every one kept.
 */
const BYTES_PER_KEPT_CHARACTER = 8;
const PART_CHARACTERS_PER_KEPT = 24;

/** 25 times what writing the files and reading them take on a developer's machine, 12 s. */
const MILLISECONDS = 300_000;

let folder: string;
beforeAll(() => {
  folder = mkdtempSync(join(tmpdir(), "xlsx-mini-read-memory-"));
});
afterAll(() => rmSync(folder, { recursive: true, force: true }));

async function written(name: string, stream: ReadableStream<Uint8Array>): Promise<string> {
  const path = join(folder, name);
  writeFileSync(path, await bytesOf(stream));
  return path;
}

describe("reading the wide table, each read in a process whose heap is held to 64 MiB", () => {
  const reads = new Map<string, FlatRead>();

  beforeAll(async () => {
    const small = await written("wide small.xlsx", streamXlsx(wideStreamed(SMALL_ROWS)));
    for (const rows of FLAT_ROWS) {
      const path = await written(`wide ${rows}.xlsx`, streamXlsx(wideStreamed(rows)));
      await Promise.all(
        SOURCES.map(async (source) => {
          const args = ["flat", path, source, small];
          reads.set(
            `${rows} ${source}`,
            (await reportOf(child, FLAT_HEAP_MEBIBYTES, args)) as FlatRead,
          );
        }),
      );
    }
  }, MILLISECONDS);

  it("runs with that limit, and reads every row from the file's bytes and from a Blob of it", () => {
    for (const rows of FLAT_ROWS) {
      for (const source of SOURCES) {
        const read = reads.get(`${rows} ${source}`)!;
        expect(read.rows, `${rows} ${source}`).toBe(rows + 1);
        expect(read.heapLimit).toBeLessThan(FLAT_HEAP_MEBIBYTES * MEBIBYTE + YOUNG_SPACE_BYTES);
      }
    }
  });

  it.each(SOURCES)("holds as much at a million rows as at 100 000, read from %s", (source) => {
    const [fewer, more] = FLAT_ROWS.map((rows) => reads.get(`${rows} ${source}`)!);
    expect(more!.peak / fewer!.peak, JSON.stringify({ fewer, more })).toBeLessThan(FLAT_RATIO);
  });
});

describe("a column kept from a sheet of 100 000 rows, each in a process whose heap is held to 256 MiB", () => {
  const columns = new Map<Kept, KeptRead>();

  beforeAll(async () => {
    await Promise.all(
      KEPT.map(async (kept) => {
        const small = await written(
          `kept ${kept} small.xlsx`,
          streamXlsx(...keptWorkbook(kept, SMALL_ROWS)),
        );
        const path = await written(
          `kept ${kept}.xlsx`,
          streamXlsx(...keptWorkbook(kept, KEPT_ROWS)),
        );
        const args = ["kept", path, kept, small];
        columns.set(kept, (await reportOf(child, KEPT_HEAP_MEBIBYTES, args)) as KeptRead);
      }),
    );
  }, MILLISECONDS);

  it.each(KEPT)("holds the text kept from %s, not the part it was read from", (kept) => {
    const column = columns.get(kept)!;
    expect(column).toMatchObject({ strings: KEPT_ROWS, wrong: 0, collected: true });
    expect(column.partBytes).toBeGreaterThan(PART_CHARACTERS_PER_KEPT * column.characters);
    expect(column.retained).toBeLessThan(BYTES_PER_KEPT_CHARACTER * column.characters);
  });
});
