import { execFileSync, execSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { delimiter, dirname, join, posix, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import config from "../vite.config.ts";
import { binary } from "./helpers/bin.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function readJson(file: string): unknown {
  return JSON.parse(readFileSync(join(root, file), "utf8"));
}

function readJsonc(file: string): unknown {
  const text = readFileSync(join(root, file), "utf8");
  return JSON.parse(text.replace(/^\s*\/\/.*$/gm, "").replace(/,(\s*[\]}])/g, "$1"));
}

const manifest = readJson("package.json") as {
  exports: Record<string, string | { types: string; default: string }>;
};

const entries: [name: string, file: string][] = Object.entries(manifest.exports).flatMap(
  ([subpath, target]) =>
    typeof target === "string"
      ? []
      : [
          [
            `xlsx-mini${subpath.slice(1)}`,
            target.default.replace(/^\.\/dist\/(.+)\.js$/, "src/$1.ts"),
          ],
        ],
);

describe("the entries", () => {
  it("are the same in the manifest, the compiler's paths, the test alias and knip", () => {
    const sorted = (pairs: string[][]) => pairs.map((pair) => pair.join(" -> ")).sort();
    const tsconfig = readJson("tsconfig.json") as {
      compilerOptions: { paths: Record<string, string[]> };
    };
    const paths = Object.entries(tsconfig.compilerOptions.paths).map(([name, files]) => [
      name,
      ...files.map((file) => posix.normalize(file)),
    ]);
    const alias = config.resolve?.alias as readonly { find: string; replacement: string }[];
    const aliased = alias.map(({ find, replacement }) => [
      find,
      relative(root, replacement).replaceAll("\\", "/"),
    ]);
    const knip = readJsonc("knip.jsonc") as { entry: string[] };
    const knipEntries = knip.entry.filter((entry) => /^src\/[^/*]+\.ts$/.test(entry));

    expect(sorted(paths)).toEqual(sorted(entries));
    expect(sorted(aliased)).toEqual(sorted(entries));
    expect(knipEntries.sort()).toEqual(entries.map(([, file]) => file).sort());
  });
});

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

  it("has every entry, as code and as declarations", () => {
    for (const target of Object.values(manifest.exports)) {
      if (typeof target === "string") continue;
      for (const file of [target.default, target.types]) {
        expect([...files.keys()]).toContain(file.replace(/^\.\/dist\//, ""));
      }
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
