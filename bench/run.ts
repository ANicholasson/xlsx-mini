// Times the three suites against the built package on whichever runtime runs this file: Node, Bun
// or Deno. With `--base <dist>` a second build is timed in the same process, call by call in turns
// with the first, so both see the same machine and the same moment; the report compares the two.
// `pnpm bench` prints a table; `--json <file>` also writes the results.
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import type * as Api from "../src/index.ts";
import type { BenchmarkResult, RunResult, Timing, Variant } from "./report.ts";
import { suites } from "./workbooks.ts";

type Write = () => Promise<Uint8Array>;

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

async function time(write: Write): Promise<number> {
  const started = performance.now();
  await write();
  return performance.now() - started;
}

function summarize(samples: readonly number[]): Timing {
  const sorted = [...samples].sort((a, b) => a - b);
  return { minMs: sorted[0]!, p50Ms: sorted[Math.floor(sorted.length / 2)]! };
}

const headDist = argument("--dist") ?? fileURLToPath(new URL("../dist/", import.meta.url));
const baseDist = argument("--base");
const head = await load(headDist);
const base = baseDist === undefined ? undefined : await load(baseDist);
const iterations = Number(argument("--iterations") ?? 7);

const benchmarks: BenchmarkResult[] = [];
for (const [suite, build] of suites) {
  const workbook = build();
  for (const variant of ["deflate", "store"] as const satisfies readonly Variant[]) {
    const options = variant === "store" ? { compression: "none" as const } : undefined;
    const headWrite: Write = () => head.writeXlsx(workbook, options);
    const baseWrite: Write | undefined =
      base === undefined ? undefined : () => base.writeXlsx(workbook, options);
    // One warm-up each, then the two builds in turns: whatever the machine does meanwhile hits both.
    await headWrite();
    await baseWrite?.();
    const headSamples: number[] = [];
    const baseSamples: number[] = [];
    for (let i = 0; i < iterations; i++) {
      headSamples.push(await time(headWrite));
      if (baseWrite !== undefined) baseSamples.push(await time(baseWrite));
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
