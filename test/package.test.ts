import { execFileSync, execSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { delimiter, dirname, join, posix } from "node:path";
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
    // Through a shell, as on Windows, an unquoted folder with a space in it is split in two.
    const outDir = tsc.shell ? `"${dist}"` : dist;
    execFileSync(tsc.file, ["-p", "tsconfig.build.json", "--outDir", outDir], {
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

describe("the build script", () => {
  it("empties dist first, so a pack never carries the files of an earlier build", () => {
    // A project of its own under the real script: the runtime tests build into this one's dist.
    const project = mkdtempSync(join(tmpdir(), "xlsx-mini-build-"));
    try {
      mkdirSync(join(project, "src"));
      mkdirSync(join(project, "dist", "old"), { recursive: true });
      writeFileSync(join(project, "src", "index.ts"), "export const built = true;\n");
      writeFileSync(join(project, "dist", "old", "stale.js.map"), "{}");
      writeFileSync(
        join(project, "tsconfig.build.json"),
        JSON.stringify({
          compilerOptions: { rootDir: "src", outDir: "dist", module: "nodenext", types: [] },
          include: ["src"],
        }),
      );
      const manifest = JSON.parse(readFileSync(join(root, "package.json"), "utf8")) as {
        scripts: { build: string };
      };
      // The script finds the compiler as a package script does: on the path.
      const path = Object.keys(process.env).find((key) => key.toLowerCase() === "path") ?? "PATH";
      const bin = join(root, "node_modules", ".bin");
      execSync(manifest.scripts.build, {
        cwd: project,
        stdio: "inherit",
        env: { ...process.env, [path]: `${bin}${delimiter}${process.env[path] ?? ""}` },
      });
      expect(readdirSync(join(project, "dist"))).toEqual(["index.js"]);
    } finally {
      rmSync(project, { recursive: true, force: true });
    }
  });
});
