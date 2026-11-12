import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, posix } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import { binary } from "./helpers/bin.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

/**
 * What the build emits is what the package carries under `dist/`. A file in it that names
 * another file, an import or a source map, must find that file in the package: a consumer's
 * bundler, editor or debugger follows the name and has nothing else to look in.
 */
describe("the built package", () => {
  const dist = mkdtempSync(join(tmpdir(), "xlsx-mini-package-"));
  /** Every emitted file, by its path under the build folder, with forward slashes. */
  let files: ReadonlyMap<string, string>;

  beforeAll(() => {
    // A build of its own, so that no other test's build is read half written.
    const tsc = binary(root, "tsc");
    execFileSync(tsc.file, ["-p", "tsconfig.build.json", "--outDir", dist], {
      cwd: root,
      stdio: "inherit",
      shell: tsc.shell,
    });
    files = new Map(
      readdirSync(dist, { recursive: true, withFileTypes: true })
        .filter((entry) => entry.isFile())
        .map((entry) => {
          const file = join(entry.parentPath, entry.name);
          const name = file.slice(dist.length + 1).replaceAll("\\", "/");
          return [name, readFileSync(file, "utf8")] as const;
        }),
    );
  });

  afterAll(() => rmSync(dist, { recursive: true, force: true }));

  it("has both entries, as code and as declarations", () => {
    for (const name of ["index.js", "index.d.ts", "browser.js", "browser.d.ts"]) {
      expect([...files.keys()]).toContain(name);
    }
  });

  it("carries no source map: the sources a map names are not in the package", () => {
    expect([...files.keys()].filter((name) => name.endsWith(".map"))).toEqual([]);
    const mapped = [...files].filter(([, text]) => text.includes("sourceMappingURL"));
    expect(mapped.map(([name]) => name)).toEqual([]);
  });

  it("names no file it does not contain", () => {
    const missing: string[] = [];
    for (const [name, text] of files) {
      // Relative specifiers of static and dynamic imports and of re-exports.
      for (const match of text.matchAll(/(?:from|import)\s*\(?\s*["'](\.{1,2}\/[^"']+)["']/g)) {
        const target = posix.join(posix.dirname(name), match[1]!);
        if (!files.has(target)) missing.push(`${name} -> ${match[1]!}`);
      }
    }
    expect(missing).toEqual([]);
  });
});
