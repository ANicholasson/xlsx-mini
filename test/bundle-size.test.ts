import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vite-plus/test";
import { type PackageSize, SIZE_BUDGET, measurePackage } from "../bench/size.ts";
import { binary } from "./helpers/bin.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

describe("bundle size", () => {
  let size: PackageSize;

  beforeAll(async () => {
    // A build of its own, so that no other test's build is read half written.
    const dist = mkdtempSync(join(tmpdir(), "xlsx-mini-size-"));
    try {
      const tsc = binary(root, "tsc");
      execFileSync(tsc.file, ["-p", "tsconfig.build.json", "--outDir", dist], {
        cwd: root,
        stdio: "inherit",
        shell: tsc.shell,
      });
      size = await measurePackage(dist);
    } finally {
      rmSync(dist, { recursive: true, force: true });
    }
  });

  it("keeps an application that imports only writeXlsx inside its budget", () => {
    expect(size.writeXlsx.gzip).toBeLessThanOrEqual(SIZE_BUDGET.writeXlsx);
  });

  it("keeps every export together inside its budget", () => {
    expect(size.everything.gzip).toBeLessThanOrEqual(SIZE_BUDGET.everything);
  });

  it("leaves out of a bundle what the application does not import", () => {
    expect(size.writeXlsx.gzip).toBeLessThan(size.everything.gzip);
    expect(size.writeXlsx.minified).toBeLessThan(size.everything.minified);
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
