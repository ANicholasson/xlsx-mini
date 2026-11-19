import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { type PackageSize, SIZE_BUDGET, measurePackage } from "../bench/size.ts";
import { binary } from "./helpers/bin.js";
import { readSideSources } from "./helpers/read-side.ts";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

const DIGESTS = join(root, "bench", "writer-digests.json");

function writerDigests(size: PackageSize): Record<string, string> {
  return {
    writeXlsx: size.writeXlsx.digest,
    streamXlsx: size.streamXlsx.digest,
    everything: size.everything.digest,
    browser: size.browser.digest,
  };
}

function modules(dist: string): string[] {
  return readdirSync(dist, { recursive: true, encoding: "utf8" })
    .map((name) => name.replaceAll("\\", "/"))
    .filter((name) => name.endsWith(".js"));
}

function eraseReadSide(dist: string): void {
  for (const entry of readdirSync(dist, { recursive: true, withFileTypes: true })) {
    const file = join(entry.parentPath, entry.name);
    const name = relative(dist, file).replaceAll("\\", "/");
    if (/^read\.|^[^/]+\/read[./]/.test(name)) rmSync(file, { recursive: true, force: true });
  }
}

describe("bundle size", () => {
  let size: PackageSize;
  let erased: PackageSize;
  let erasedModules: string[];

  beforeAll(async () => {
    // A build of its own, so that no other test's build is read half written.
    const dist = mkdtempSync(join(tmpdir(), "xlsx-mini-size-"));
    try {
      const tsc = binary(root, "tsc");
      // Through a shell, as on Windows, an unquoted folder with a space in it is split in two.
      const outDir = tsc.shell ? `"${dist}"` : dist;
      execFileSync(tsc.file, ["-p", "tsconfig.build.json", "--outDir", outDir], {
        cwd: root,
        stdio: "inherit",
        shell: tsc.shell,
      });
      size = await measurePackage(dist);
      const built = modules(dist);
      eraseReadSide(dist);
      erasedModules = built.filter((module) => !existsSync(join(dist, module))).sort();
      erased = await measurePackage(dist);
    } finally {
      rmSync(dist, { recursive: true, force: true });
    }
    if (process.env["UPDATE_DIGESTS"] === "1") {
      writeFileSync(DIGESTS, `${JSON.stringify(writerDigests(size), null, 2)}\n`);
    }
  });

  it("keeps an application that imports only writeXlsx inside its budget", () => {
    expect(size.writeXlsx.gzip).toBeLessThanOrEqual(SIZE_BUDGET.writeXlsx);
  });

  it("keeps every export together inside its budget", () => {
    expect(size.everything.gzip).toBeLessThanOrEqual(SIZE_BUDGET.everything);
  });

  it("keeps the reader inside its budget", () => {
    expect(size.read?.gzip).toBeLessThanOrEqual(SIZE_BUDGET.read);
  });

  it("leaves out of a bundle what the application does not import", () => {
    expect(size.writeXlsx.gzip).toBeLessThan(size.everything.gzip);
    expect(size.writeXlsx.minified).toBeLessThan(size.everything.minified);
  });

  it("bundles the writer to the same bytes with the read side erased from the build", () => {
    const readSide = readSideSources(root).map((file) => file.replace(/^src\/(.+)\.ts$/, "$1.js"));
    expect(erasedModules, "erased the read side's modules").toEqual(readSide);
    expect(erased.read).toBeUndefined();
    expect(writerDigests(erased)).toEqual(writerDigests(size));
  });

  it("bundles the writer to the bytes of its committed digests", () => {
    const committed: unknown = JSON.parse(readFileSync(DIGESTS, "utf8"));
    expect(
      writerDigests(size),
      "The writer's bytes changed. If the change means to, `pnpm digests:update` records it.",
    ).toEqual(committed);
  });

  it("measures what a bundler does with the package: no module has side effects", () => {
    const manifest = JSON.parse(readFileSync(join(root, "package.json"), "utf8")) as {
      sideEffects?: boolean;
    };
    expect(manifest.sideEffects).toBe(false);
  });

  it("is stated in the README as the budgets it is held to", () => {
    const readme = readFileSync(join(root, "README.md"), "utf8");
    expect(readme).toContain(`under ${SIZE_BUDGET.writeXlsx / 1000} kB`);
    expect(readme).toContain(`under ${SIZE_BUDGET.everything / 1000} kB`);
  });
});
