import { execFileSync, spawnSync } from "node:child_process";
import { cpSync, mkdtempSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, describe, expect, it } from "vite-plus/test";
import { binary } from "./helpers/bin.js";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");

function git(dir: string, ...args: string[]): void {
  execFileSync("git", args, { cwd: dir, stdio: "pipe" });
}

/** Runs the hook's task dispatcher on what is staged in `dir`, as the pre-commit hook does. */
function staged(dir: string) {
  const vp = binary(root, "vp");
  return spawnSync(vp.file, ["staged"], { cwd: dir, encoding: "utf8", shell: vp.shell });
}

/** A throwaway repository with this project's configuration and dependencies, nothing else. */
function scratchRepository(): string {
  const dir = mkdtempSync(join(tmpdir(), "xlsx-mini-hook-"));
  for (const file of ["package.json", "vite.config.ts"]) cpSync(join(root, file), join(dir, file));
  writeFileSync(
    join(dir, "tsconfig.json"),
    JSON.stringify({ compilerOptions: { strict: true, module: "ESNext" }, include: ["*.ts"] }),
  );
  symlinkSync(join(root, "node_modules"), join(dir, "node_modules"), "junction");
  git(dir, "init", "-q");
  git(dir, "config", "user.email", "hook@example.invalid");
  git(dir, "config", "user.name", "hook test");
  git(dir, "add", "package.json", "vite.config.ts", "tsconfig.json");
  git(dir, "commit", "-q", "-m", "configuration");
  return dir;
}

describe("pre-commit staged tasks", () => {
  let dir: string;
  beforeAll(() => {
    dir = scratchRepository();
  }, 30_000);
  afterAll(() => rmSync(dir, { recursive: true, force: true }));

  // Regression test from the Codex review of pull request #15: the linter used to run on commits
  // that staged only JSON, Markdown or YAML, found nothing to lint, and blocked the commit.
  it("lets a commit through that stages only JSON, Markdown or YAML", () => {
    writeFileSync(join(dir, "notes.md"), "# Notes\n");
    writeFileSync(join(dir, "data.json"), "{}\n");
    git(dir, "add", "notes.md", "data.json");
    const result = staged(dir);
    expect(result.status, result.stdout + result.stderr).toBe(0);
    git(dir, "commit", "-q", "-m", "documents");
  }, 60_000);

  it("still blocks a commit that stages code the linter rejects", () => {
    writeFileSync(join(dir, "bad.ts"), "export const loose: any = 1;\n");
    git(dir, "add", "bad.ts");
    const result = staged(dir);
    expect(result.status, result.stdout + result.stderr).not.toBe(0);
  }, 60_000);
});
