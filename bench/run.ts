// Times the three suites against the built package on whichever runtime runs this file: Node, Bun
// or Deno. With `--base <dist>` a second build is timed in the same process, call by call in turns
// with the first, so both see the same machine and the same moment; the report compares the two.
// Nothing is written before the suites and nothing is done to one build that is not done to the
// other: the first write each build makes is a large one, as a first export in a fresh process is,
// and a small write that only one of them made first once showed as fifteen percent between two
// copies of the same build. `pnpm bench` prints a table; `--json <file>` also writes the results.
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import type * as Api from "../src/index.ts";
import type { BenchmarkResult, RunResult, Timing, Variant } from "./report.ts";
import { suites, withAutoWidths } from "./workbooks.ts";

type Write = () => Promise<unknown>;

interface RuntimeInfo {
  readonly runtime: string;
  readonly version: string;
}

/** Which engine is running this file, by the global each one exposes. */
function detectRuntime(): RuntimeInfo {
  const host = globalThis as {
    Bun?: { version: string };
    Deno?: { version: { deno: string } };
    process?: { versions: { node: string } };
  };
  if (host.Bun !== undefined) return { runtime: "bun", version: host.Bun.version };
  if (host.Deno !== undefined) return { runtime: "deno", version: host.Deno.version.deno };
  return { runtime: "node", version: host.process?.versions.node ?? "unknown" };
}

function argument(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  return index === -1 ? undefined : process.argv[index + 1];
}

/** A build of the package, loaded by the host itself from its `dist/` folder. */
async function load(dist: string): Promise<typeof Api> {
  return (await import(pathToFileURL(resolve(dist, "index.js")).href)) as typeof Api;
}

/** Node with `--expose-gc` and Bun expose a collector; Deno does not, and then nothing happens. */
function collectGarbage(): void {
  const host = globalThis as { gc?: () => void; Bun?: { gc?: (force: boolean) => void } };
  if (host.gc !== undefined) host.gc();
  else host.Bun?.gc?.(true);
}

async function time(write: Write): Promise<number> {
  const started = performance.now();
  await write();
  return performance.now() - started;
}

function summarize(samples: readonly number[]): Timing {
  const sorted = [...samples].sort((a, b) => a - b);
  return { minMs: sorted[0]!, p50Ms: sorted[Math.floor(sorted.length / 2)]! };
}

/** Whether a write completes: a base build from before a feature refuses or lacks its variant. */
async function succeeds(write: Write): Promise<boolean> {
  try {
    await write();
    return true;
  } catch {
    return false;
  }
}

const headDist = argument("--dist") ?? fileURLToPath(new URL("../dist/", import.meta.url));
const baseDist = argument("--base");
const head = await load(headDist);
const base = baseDist === undefined ? undefined : await load(baseDist);
const iterations = Number(argument("--iterations") ?? 7);

const benchmarks: BenchmarkResult[] = [];
for (const [suite, build] of suites) {
  const workbook = build();
  const autoWorkbook = withAutoWidths(workbook);
  for (const variant of [
    "deflate",
    "store",
    "stream",
    "auto",
  ] as const satisfies readonly Variant[]) {
    const store = { compression: "none" as const };
    const writeWith = (api: typeof Api): Write => {
      if (variant === "deflate") return () => api.writeXlsx(workbook);
      if (variant === "store") return () => api.writeXlsx(workbook, store);
      if (variant === "auto") return () => api.writeXlsx(autoWorkbook, store);
      return () => api.streamXlsx(workbook).pipeTo(new WritableStream());
    };
    const headWrite = writeWith(head);
    let baseWrite = base === undefined ? undefined : writeWith(base);
    // One warm-up each, then the two builds in turns, the order swapped every iteration so that
    // neither side always pays for the other's garbage or always gets the warmer cache; a
    // collection before each timed call, where the host exposes one, takes the rest of that bias.
    await headWrite();
    // A base from before streaming or auto widths fails its warm-up, and the row then shows no
    // base rather than a wrong one. The build under test must write every variant.
    if (baseWrite !== undefined && !(await succeeds(baseWrite))) baseWrite = undefined;
    const headSamples: number[] = [];
    const baseSamples: number[] = [];
    for (let i = 0; i < iterations; i++) {
      const order =
        baseWrite !== undefined && i % 2 === 1 ? [baseWrite, headWrite] : [headWrite, baseWrite];
      for (const write of order) {
        if (write === undefined) continue;
        collectGarbage();
        const elapsed = await time(write);
        (write === headWrite ? headSamples : baseSamples).push(elapsed);
      }
    }
    const result: BenchmarkResult = {
      suite,
      variant,
      samples: iterations,
      head: summarize(headSamples),
      ...(baseWrite === undefined ? {} : { base: summarize(baseSamples) }),
    };
    benchmarks.push(result);
  }
}

const run: RunResult = {
  ...detectRuntime(),
  arch: process.arch,
  commit: argument("--commit") ?? process.env["GITHUB_SHA"] ?? "local",
  ...(baseDist === undefined ? {} : { baseCommit: argument("--base-commit") ?? "base" }),
  date: new Date().toISOString(),
  benchmarks,
};

console.log(
  `${run.runtime} ${run.version} on ${run.arch}, ${iterations} runs per benchmark${base === undefined ? "" : ", in turns with the base build"}`,
);
console.table(
  benchmarks.map((b) => ({
    suite: b.suite,
    variant: b.variant,
    "min ms": b.head.minMs.toFixed(1),
    "p50 ms": b.head.p50Ms.toFixed(1),
    ...(b.base === undefined ? {} : { "base min ms": b.base.minMs.toFixed(1) }),
  })),
);

const out = argument("--json");
if (out !== undefined) {
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, `${JSON.stringify(run, null, 2)}\n`);
  console.log(`written ${out}`);
}
