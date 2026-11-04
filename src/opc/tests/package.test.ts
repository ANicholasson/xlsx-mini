import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vite-plus/test";
import { unzip } from "../../../test/helpers/unzip.js";
import { createPackage } from "../index.js";

const encode = (text: string) => new TextEncoder().encode(text);

describe("createPackage", () => {
  const parts = [
    { path: "[Content_Types].xml", data: encode("<Types/>") },
    { path: "xl/worksheets/sheet1.xml", data: encode("<worksheet>æøå</worksheet>") },
    { path: "empty.txt", data: new Uint8Array(0) },
  ];

  it("writes STORE entries that an independent reader walks back in order", () => {
    const entries = unzip(createPackage(parts));
    expect(entries.map((e) => e.name)).toEqual(parts.map((p) => p.path));
    expect(entries.every((e) => e.method === 0)).toBe(true);
    expect(entries.map((e) => new TextDecoder().decode(e.data))).toEqual([
      "<Types/>",
      "<worksheet>æøå</worksheet>",
      "",
    ]);
    expect(entries[1]?.uncompressedSize).toBe(parts[1]?.data.byteLength);
    expect(entries[0]?.localHeaderOffset).toBe(0);
  });

  it("is byte-identical across calls", () => {
    expect(createPackage(parts)).toEqual(createPackage(parts));
  });

  it("passes the system unzip integrity test", () => {
    const dir = mkdtempSync(join(tmpdir(), "mini-xlsx-"));
    const file = join(dir, "package.zip");
    writeFileSync(file, createPackage(parts));
    const report = execFileSync("unzip", ["-t", file], { encoding: "utf8" });
    expect(report).toContain("No errors detected");
  });
});
