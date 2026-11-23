import { execFile } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { rolldown } from "rolldown";

/**
 * Runs `child`, a TypeScript file, in a process of its own whose heap is held to `heapMebibytes`
 * and which has a collector to call, and parses the line of JSON it prints. The process fails, and
 * so does the promise, when the heap does not hold what the child reads.
 */
export async function reportOf(child: string, heapMebibytes: number): Promise<unknown> {
  const directory = mkdtempSync(join(tmpdir(), "xlsx-mini-child-"));
  try {
    // Node 20 does not run TypeScript: bundled, the child is one file it does.
    const build = await rolldown({ input: child, platform: "node", logLevel: "silent" });
    const { output } = await build.generate({ format: "esm" });
    await build.close();
    const script = join(directory, "child.mjs");
    writeFileSync(script, output[0].code);
    const { stdout } = await promisify(execFile)(
      process.execPath,
      [`--max-old-space-size=${heapMebibytes}`, "--expose-gc", script],
      { encoding: "utf8" },
    );
    return JSON.parse(stdout);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}
