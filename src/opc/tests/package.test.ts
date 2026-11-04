import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { unzip } from "../../../test/helpers/unzip.js";
import { createPackage, selectCompression, store } from "../index.js";

const encode = (text: string) => new TextEncoder().encode(text);

const parts = [
  { path: "[Content_Types].xml", data: encode("<Types/>") },
  {
    path: "xl/worksheets/sheet1.xml",
    data: encode(`<worksheet>${"æøå ".repeat(500)}</worksheet>`),
  },
  { path: "empty.txt", data: new Uint8Array(0) },
];

function passesSystemUnzip(bytes: Uint8Array): boolean {
  const dir = mkdtempSync(join(tmpdir(), "xlsx-mini-"));
  const file = join(dir, "package.zip");
  writeFileSync(file, bytes);
  return execFileSync("unzip", ["-t", file], { encoding: "utf8" }).includes("No errors detected");
}

describe("createPackage with STORE", () => {
  it("writes entries an independent reader walks back in order", async () => {
    const entries = unzip(await createPackage(parts, store));
    expect(entries.map((e) => e.name)).toEqual(parts.map((p) => p.path));
    expect(entries.every((e) => e.method === 0)).toBe(true);
    expect(entries.map((e) => new TextDecoder().decode(e.data))).toEqual(
      parts.map((p) => new TextDecoder().decode(p.data)),
    );
    expect(entries[0]?.localHeaderOffset).toBe(0);
  });

  it("is byte-identical across calls", async () => {
    expect(await createPackage(parts, store)).toEqual(await createPackage(parts, store));
  });

  it("passes the system unzip integrity test", async () => {
    expect(passesSystemUnzip(await createPackage(parts, store))).toBe(true);
  });
});

describe("createPackage with DEFLATE", () => {
  it("writes method-8 entries that inflate back to the original parts", async () => {
    const entries = unzip(await createPackage(parts, selectCompression("deflate")));
    expect(entries.map((e) => e.method)).toEqual([8, 8, 8]);
    expect(entries.map((e) => new TextDecoder().decode(e.data))).toEqual(
      parts.map((p) => new TextDecoder().decode(p.data)),
    );
    expect(entries[1]!.compressedSize).toBeLessThan(entries[1]!.uncompressedSize / 4);
  });

  it("passes the system unzip integrity test, which inflates independently", async () => {
    expect(passesSystemUnzip(await createPackage(parts, selectCompression("deflate")))).toBe(true);
  });
});

describe("createPackage limits", () => {
  it("refuses a package that would exceed 4 GiB before allocating it", async () => {
    const huge = {
      method: 8 as const,
      compress: () => Promise.resolve({ length: 2 ** 32 } as unknown as Uint8Array),
    };
    await expect(createPackage(parts, huge)).rejects.toMatchObject({
      code: "PACKAGE_TOO_LARGE",
      path: "workbook",
    });
  });

  it("gathers a multi-chunk deflate stream back into one payload", async () => {
    const big = {
      path: "big.txt",
      data: encode(
        Array.from({ length: 200_000 }, (_, i) => `row ${i} ${(i * 7919) % 1000}`).join("\n"),
      ),
    };
    const entries = unzip(await createPackage([big], selectCompression("deflate")));
    expect(entries[0]!.uncompressedSize).toBe(big.data.byteLength);
    expect(new TextDecoder().decode(entries[0]!.data)).toBe(new TextDecoder().decode(big.data));
  });
});

describe("selectCompression", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("prefers deflate when the runtime supports deflate-raw", () => {
    expect(selectCompression("deflate").method).toBe(8);
    expect(selectCompression("none").method).toBe(0);
  });

  it("falls back to STORE when CompressionStream is missing", () => {
    vi.stubGlobal("CompressionStream", undefined);
    expect(selectCompression("deflate").method).toBe(0);
  });

  it("falls back to STORE when the runtime lists deflate-raw but refuses it", () => {
    vi.stubGlobal(
      "CompressionStream",
      class {
        constructor(format: string) {
          throw new TypeError(`unsupported format ${format}`);
        }
      },
    );
    expect(selectCompression("deflate").method).toBe(0);
  });
});
