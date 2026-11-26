import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import { type Bundled, bundled, supervise } from "../helpers/child.ts";
import { unzip } from "../helpers/unzip.ts";
import { MAIN, RELATIONSHIPS } from "./book.ts";
import { DEFAULT_HEAP_MEBIBYTES, DEFAULT_MILLISECONDS, hostileCases } from "./cases.ts";
import type { HostileMessage } from "./hostile.child.ts";

const here = dirname(fileURLToPath(import.meta.url));

const MEBIBYTE = 1_048_576;
/**
 * What a heap may hold besides its old space, which the cap limits: the young space, 96 MiB on
 * Node 26 and 192 MiB on Node 24 on the machine that set this. Without the cap the heap is
 * gibibytes.
 */
const YOUNG_SPACE_BYTES = 256 * MEBIBYTE;

let child: Bundled;
beforeAll(async () => {
  child = await bundled(join(here, "hostile.child.ts"));
}, 60_000);
afterAll(() => child.remove());

describe("a hostile file, read in a process of its own", () => {
  it("names each case once", () => {
    const names = hostileCases.map(({ name }) => name);
    expect(new Set(names).size).toBe(names.length);
  });

  it.concurrent.each(hostileCases.map((hostile) => [hostile.name, hostile] as const))(
    "%s",
    async (name, hostile) => {
      const heapMebibytes = hostile.heapMebibytes ?? DEFAULT_HEAP_MEBIBYTES;
      const milliseconds = hostile.milliseconds ?? DEFAULT_MILLISECONDS;
      const { messages, exit, overran, stderr } = await supervise<HostileMessage>(
        child.script,
        [name],
        heapMebibytes,
        ({ kind }) => (kind === "built" ? milliseconds : undefined),
      );
      expect(overran, `${name}: passed its time bound of ${milliseconds} ms`).toBeUndefined();
      const ended = messages.at(-1);
      const why = ended?.kind === "failed" ? ended.failure : stderr;
      expect(ended?.kind, `${name}: ended (${exit}) without an outcome: ${why}`).toBe("read");
      const built = messages.find((message) => message.kind === "built");
      expect(built?.heapLimit, `${name}: the heap cap took effect`).toBeLessThan(
        heapMebibytes * MEBIBYTE + YOUNG_SPACE_BYTES,
      );
      if (ended?.kind !== "read") return;
      expect(ended.same, `${name}\nexpected ${ended.expected}\nread     ${ended.outcome}`).toBe(
        true,
      );
    },
    120_000,
  );
});

function fileOf(name: string): Uint8Array {
  const file = hostileCases.find((hostile) => hostile.name === name)?.file();
  if (!(file instanceof Uint8Array)) throw new Error(`no case ${name} of bytes`);
  return file;
}

const textOf = (file: Uint8Array, path: string) =>
  new TextDecoder().decode(unzip(file).find((entry) => entry.name === path)?.data);

describe("the files the issue names, as the list builds them", () => {
  it("a damaged sheet relationship type differs from a worksheet's by one character", () => {
    const type = /Type="([^"]*)" Target="worksheets\/sheet1.xml"/.exec(
      textOf(fileOf("a damaged sheet relationship type"), "xl/_rels/workbook.xml.rels"),
    )?.[1];
    const worksheet = `${RELATIONSHIPS}/worksheet`;
    expect(type).toHaveLength(worksheet.length);
    const differing = Array.from({ length: worksheet.length }, (_, at) => at).filter(
      (at) => worksheet[at] !== type?.[at],
    );
    expect(differing).toHaveLength(1);
  });

  it("a directory of four million entries is one, in a package of nothing else", () => {
    const file = fileOf("a directory of four million entries");
    const view = new DataView(file.buffer, file.byteOffset, file.byteLength);
    const zip64End = file.length - 22 - 20 - 56;
    expect(view.getUint32(zip64End, true)).toBe(0x06064b50);
    expect(view.getBigUint64(zip64End + 32, true)).toBe(4_000_000n);
    expect(view.getBigUint64(zip64End + 40, true)).toBe(4_000_000n * 53n);
  });

  it.each(["cells", "runs"])(
    "three sheets of one row of 119 MiB in %s name one part, in a file under 2 MiB",
    (row) => {
      const file = fileOf(`three sheets of one row of 119 MiB in ${row}, read at once`);
      const sheet = unzip(file).find((entry) => entry.name === "xl/worksheets/sheet1.xml");
      expect(sheet?.uncompressedSize).toBeGreaterThan(118 * MEBIBYTE);
      expect(sheet?.uncompressedSize).toBeLessThanOrEqual(128 * MEBIBYTE);
      expect(file.length).toBeLessThan(2 * MEBIBYTE);
      expect(textOf(file, "xl/workbook.xml").match(/r:id="rId1"/g)).toHaveLength(3);
    },
  );

  it("a chart sheet stands beside a string table over the limit it is read with", () => {
    const name = "a chart sheet beside an over-limit string table";
    const file = fileOf(name);
    const limit = hostileCases.find((hostile) => hostile.name === name)?.limits?.sharedStringBytes;
    expect(textOf(file, "xl/sharedStrings.xml").length).toBeGreaterThan(limit ?? Infinity);
    expect(textOf(file, "xl/chartsheets/sheet1.xml")).toBe(`<chartsheet xmlns="${MAIN}"/>`);
  });
});
