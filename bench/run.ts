// Times the three suites against the built package, on whichever runtime runs this file: Node, Bun
// or Deno. `pnpm bench` prints a table; `--json <file>` also writes the results for the report.
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { Bench } from "tinybench";
import type * as Api from "../src/index.ts";
import type { BenchmarkResult, RunResult, Variant } from "./report.ts";
import { suites } from "./workbooks.ts";

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

// The built package, what users run, loaded by the host itself; `pnpm bench` builds it first.
// The specifier is computed so that no tool resolves `dist/` before a build exists.
const { writeXlsx } = (await import(
  new URL("../dist/index.js", import.meta.url).href
)) as typeof Api;

const iterations = Number(argument("--iterations") ?? 7);
const bench = new Bench({ iterations, time: 0, warmupIterations: 1, warmupTime: 0 });
const labels: Array<{ suite: string; variant: Variant }> = [];
for (const [name, build] of suites) {
  const workbook = build();
  bench.add(`${name}, deflate`, async () => {
    await writeXlsx(workbook);
  });
  labels.push({ suite: name, variant: "deflate" });
  bench.add(`${name}, store`, async () => {
    await writeXlsx(workbook, { compression: "none" });
  });
  labels.push({ suite: name, variant: "store" });
}
await bench.run();

const benchmarks: BenchmarkResult[] = bench.tasks.map((task, index) => {
  const result = task.result;
  if (result === undefined || !("latency" in result) || result.latency === undefined) {
    throw new Error(`${task.name}: no result`);
  }
  const { suite, variant } = labels[index]!;
  const latency = result.latency;
  return {
    suite,
    variant,
    samples: latency.samplesCount,
    p50Ms: latency.p50,
    minMs: latency.min,
    meanMs: latency.mean,
  };
});

const run: RunResult = {
  ...detectRuntime(),
  arch: process.arch,
  commit: argument("--commit") ?? process.env["GITHUB_SHA"] ?? "local",
  date: new Date().toISOString(),
  benchmarks,
};

console.log(`${run.runtime} ${run.version} on ${run.arch}, ${iterations} iterations per benchmark`);
console.table(
  benchmarks.map((b) => ({
    suite: b.suite,
    variant: b.variant,
    "p50 ms": b.p50Ms.toFixed(1),
    "min ms": b.minMs.toFixed(1),
    "mean ms": b.meanMs.toFixed(1),
  })),
);

const out = argument("--json");
if (out !== undefined) {
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, `${JSON.stringify(run, null, 2)}\n`);
  console.log(`written ${out}`);
}
