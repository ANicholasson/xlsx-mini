import { spawnSync } from "node:child_process";
import { cpSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import { binary } from "./helpers/bin.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

const PLANTED: Readonly<Record<string, readonly string[]>> = {
  "src/read.ts": ["./errors/read.ts", "./opc/read.ts"],
  "src/errors/read.ts": ["./index.ts"],
  "src/errors/index.ts": [],
  "src/write-xlsx.ts": ["./errors/read.ts", "./opc/reader.ts"],
  "src/opc/index.ts": ["./read.ts", "./lib/zip.ts"],
  "src/opc/reader.ts": [],
  "src/opc/read.ts": ["./read/part.ts", "../errors/index.ts"],
  "src/opc/read/part.ts": ["../lib/crc32.ts"],
  "src/opc/lib/crc32.ts": [],
  "src/opc/lib/zip.ts": ["./crc32.ts", "../read/part.ts"],
  "src/opc/tests/read.test.ts": ["../read.ts"],
  "test/read.test.ts": ["../src/read.ts"],
  "bench/read.ts": ["../src/read.ts"],
};

interface Violation {
  readonly from: string;
  readonly to: string;
  readonly rule: { readonly name: string };
}

describe("the boundary between the write side and the read side", () => {
  let dir: string;
  let violations: readonly Violation[];

  beforeAll(() => {
    dir = mkdtempSync(join(tmpdir(), "xlsx-mini-boundaries-"));
    cpSync(join(root, ".dependency-cruiser.cjs"), join(dir, ".dependency-cruiser.cjs"));
    writeFileSync(join(dir, "tsconfig.json"), "{}\n");
    for (const [file, imports] of Object.entries(PLANTED)) {
      mkdirSync(dirname(join(dir, file)), { recursive: true });
      writeFileSync(
        join(dir, file),
        `${imports.map((target) => `import "${target}";\n`).join("")}export {};\n`,
      );
    }
    const depcruise = binary(root, "depcruise");
    const result = spawnSync(depcruise.file, ["src", "test", "bench", "--output-type", "json"], {
      cwd: dir,
      encoding: "utf8",
      shell: depcruise.shell,
    });
    violations = (JSON.parse(result.stdout) as { summary: { violations: Violation[] } }).summary
      .violations;
  }, 30_000);

  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  it("refuses an import into the read side from a writer file, and allows tests and benchmarks", () => {
    expect(violations.map(({ rule, from, to }) => `${rule.name}: ${from} -> ${to}`).sort()).toEqual(
      [
        "write-side-never-imports-read-side: src/opc/index.ts -> src/opc/read.ts",
        "write-side-never-imports-read-side: src/opc/lib/zip.ts -> src/opc/read/part.ts",
        "write-side-never-imports-read-side: src/write-xlsx.ts -> src/errors/read.ts",
      ],
    );
  });
});
