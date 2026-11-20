import { execFile } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { rolldown } from "rolldown";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import type { Flood, Report } from "./read-memory.child.ts";

const here = dirname(fileURLToPath(import.meta.url));

const MEBIBYTE = 1_048_576;

const HEAP_MEBIBYTES = 512;
const REFUSAL_MILLISECONDS = 2_000;
/** A token is read again once the buffer has doubled: twice the longest token, and two chunks. */
const TOKEN_FLOOD_BYTES = 2 * 4_194_304 + 2 * 16_384;
const FLOOD_HEAP_BYTES = 64 * MEBIBYTE;
const KEPT_HEAP_BYTES = 2 * MEBIBYTE;

describe("the tokenizer in a process whose heap is held to 512 MiB", () => {
  let report: Report;

  const floodsOf = (name: string): readonly Flood[] =>
    report.floods.filter((flood) => flood.name === name);

  beforeAll(async () => {
    const directory = mkdtempSync(join(tmpdir(), "xlsx-mini-xml-"));
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
        [`--max-old-space-size=${HEAP_MEBIBYTES}`, "--expose-gc", script],
        { encoding: "utf8" },
      );
      report = JSON.parse(stdout) as Report;
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  }, 120_000);

  it("runs with that limit, and every flood is as large as a part may be under the default limits", () => {
    expect(report.heapLimit).toBeLessThan(1024 * MEBIBYTE);
    expect(report.floodBytes).toBe(128 * MEBIBYTE);
  });

  it("refuses a flood of open tags at the 257th, having read one chunk", () => {
    const floods = floodsOf("open tags");
    expect(floods.map(({ chunkBytes }) => chunkBytes)).toEqual([16_384, 16 * MEBIBYTE]);
    for (const { chunkBytes, outcome, bytesRead, milliseconds, heapGrowth } of floods) {
      expect(outcome, `${chunkBytes}`).toBe("INVALID_XML");
      expect(bytesRead, `${chunkBytes}`).toBe(chunkBytes);
      expect(milliseconds, `${chunkBytes}`).toBeLessThan(REFUSAL_MILLISECONDS);
      expect(heapGrowth, `${chunkBytes}`).toBeLessThan(FLOOD_HEAP_BYTES);
    }
  });

  it("refuses a flood of one long token of every kind before a sixteenth of it is read", () => {
    const kinds = [
      "one run of text",
      "one attribute value",
      "one name",
      "white space in one tag",
      "white space in one close tag",
      "one comment",
      "one CDATA section",
      "one instruction",
      "one reference",
      "attributes of one tag",
    ];
    for (const kind of kinds) {
      const [flood] = floodsOf(kind);
      expect(flood, kind).toMatchObject({ chunkBytes: 16_384, outcome: "INVALID_XML" });
      expect(flood!.bytesRead, kind).toBeLessThanOrEqual(TOKEN_FLOOD_BYTES);
      expect(flood!.bytesRead, kind).toBeGreaterThan(4_194_304);
      expect(flood!.milliseconds, kind).toBeLessThan(REFUSAL_MILLISECONDS);
      expect(flood!.heapGrowth, kind).toBeLessThan(FLOOD_HEAP_BYTES);
    }
  });

  it("refuses the text of one element that many comments split, once it is longer than a token may be", () => {
    const [flood] = floodsOf("the text of one element, in runs between comments");
    expect(flood).toMatchObject({ outcome: "INVALID_XML" });
    expect(flood!.bytesRead).toBeGreaterThan((4_194_304 * 23) / 16);
    expect(flood!.bytesRead).toBeLessThan((4_194_304 * 23) / 16 + 2 * 16_384);
    expect(flood!.heapGrowth).toBeLessThan(FLOOD_HEAP_BYTES);
  });

  it("refuses nested elements whose names are longer than a token may be, before it holds them", () => {
    const [flood] = floodsOf("the long names of 256 nested elements");
    expect(flood).toMatchObject({ outcome: "INVALID_XML" });
    expect(flood!.bytesRead).toBeLessThan(10 * 500_010);
    expect(flood!.heapGrowth).toBeLessThan(FLOOD_HEAP_BYTES);
  });

  it("holds the text of one element that long comments split, not the chunks it was cut from", () => {
    const [flood] = floodsOf("the text of one element, in runs between long comments");
    expect(flood).toMatchObject({ outcome: "read" });
    expect(flood!.bytesRead).toBeGreaterThan(report.floodBytes);
    expect(flood!.heapGrowth).toBeLessThan(4 * MEBIBYTE);
  });

  it("refuses a long token that comes in the largest chunks an inflate gives, at the first", () => {
    for (const kind of ["one run of text", "one attribute value"]) {
      const flood = floodsOf(kind).find(({ chunkBytes }) => chunkBytes === 16 * MEBIBYTE);
      expect(flood, kind).toMatchObject({ outcome: "INVALID_XML", bytesRead: 16 * MEBIBYTE });
      expect(flood!.milliseconds, kind).toBeLessThan(REFUSAL_MILLISECONDS);
      expect(flood!.heapGrowth, kind).toBeLessThan(3 * 16 * MEBIBYTE);
    }
  });

  it("reads white space after the root to its end without holding it", () => {
    const [flood] = floodsOf("white space after the root");
    expect(flood).toMatchObject({ outcome: "read", bytesRead: 128 * MEBIBYTE });
    expect(flood!.heapGrowth).toBeLessThan(4 * MEBIBYTE);
  });

  it("holds only what a visitor kept of a part of 64 MiB: no text, value or name keeps its chunk alive", () => {
    expect(report.partBytes).toBe(64 * MEBIBYTE);
    const kept = report.kept.slice(0, 3);
    expect(kept.map(({ what }) => what)).toEqual([
      "the text of an element",
      "the value of an attribute",
      "the name of an element",
    ]);
    for (const { what, strings, characters, heapGrowth } of kept) {
      // More strings than the part has chunks, each longer than an engine copies of itself.
      expect(strings, what).toBeGreaterThan(report.partBytes / 16_384);
      expect(characters / strings, what).toBeGreaterThan(13);
      expect(heapGrowth, what).toBeLessThan(KEPT_HEAP_BYTES);
    }
  });

  it("lets go of a tag of as many attributes as a token may hold once the next tag is read", () => {
    expect(report.heldAfterAttributes).toBeLessThan(MEBIBYTE);
  });

  it("sees a chunk that is kept alive: slices of the buffer, which the contract forbids keeping, hold the part", () => {
    const sliced = report.kept[3]!;
    expect(sliced.what).toBe("a slice of the buffer, which a visitor must not keep");
    expect(sliced.heapGrowth).toBeGreaterThan(report.partBytes / 2);
  });
});
