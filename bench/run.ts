// Times the three suites against the built package on whichever runtime runs this file: Node, Bun
// or Deno. Every build is timed in processes of its own, started by this one: a process loads one
// build, writes nothing before the suites, warms each variant once and times it. With
// `--base <dist>` the processes of the two builds take turns, so both see the same machine at
// nearly the same moment, and the report compares the minima. Two builds in one process share the
// engine's compiled code and its heap, and that showed as differences between two copies of the
// same build; apart, a build meets what a program using it meets.
// `pnpm bench` prints a table; `--json <file>` also writes the results.
import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import type * as Api from "../src/index.ts";
import type { BenchmarkResult, RunResult, Timing, Variant } from "./report.ts";
import { suites, withAutoWidths } from "./workbooks.ts";

type Write = () => Promise<unknown>;

const VARIANTS = ["deflate", "store", "stream", "auto"] as const satisfies readonly Variant[];

/** What one process measured for one suite written one way; no samples when the build cannot. */
interface Measurement {
  readonly suite: string;
  readonly variant: Variant;
  readonly samples: readonly number[] | null;
}

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

/** Whether a write completes: a build from before a feature refuses or lacks its variant. */
async function succeeds(write: Write): Promise<boolean> {
  try {
    await write();
    return true;
  } catch {
    return false;
  }
}

/**
 * What a measuring process does: loads the one build from its `dist/` folder, as the host itself
 * resolves it, and for every suite and variant writes once to warm up, then `iterations` timed
 * times with a collection before each where the host exposes one.
 */
async function measure(dist: string, iterations: number): Promise<Measurement[]> {
  const api = (await import(pathToFileURL(resolve(dist, "index.js")).href)) as typeof Api;
  const store = { compression: "none" as const };
  const measurements: Measurement[] = [];
  for (const [suite, build] of suites) {
    const workbook = build();
    const autoWorkbook = withAutoWidths(workbook);
    const writes: Readonly<Record<Variant, Write>> = {
      deflate: () => api.writeXlsx(workbook),
      store: () => api.writeXlsx(workbook, store),
      stream: () => api.streamXlsx(workbook).pipeTo(new WritableStream()),
      auto: () => api.writeXlsx(autoWorkbook, store),
    };
    for (const variant of VARIANTS) {
      const write = writes[variant];
      if (!(await succeeds(write))) {
        measurements.push({ suite, variant, samples: null });
        continue;
      }
      const samples: number[] = [];
      for (let i = 0; i < iterations; i++) {
        collectGarbage();
        samples.push(await time(write));
      }
      measurements.push({ suite, variant, samples });
    }
  }
  return measurements;
}

/** Runs this file again on the same runtime, with the same engine flags, to measure one build. */
function measureInProcess(dist: string, iterations: number): Measurement[] {
  const script = fileURLToPath(import.meta.url);
  const engineArguments =
    detectRuntime().runtime === "deno" ? ["run", "--allow-all", "--no-lock"] : process.execArgv;
  const child = spawnSync(
    process.execPath,
    [...engineArguments, script, "--measure", dist, "--iterations", String(iterations)],
    { encoding: "utf8", stdio: ["ignore", "pipe", "inherit"], maxBuffer: 64 * 1024 * 1024 },
  );
  if (child.status !== 0) {
    throw new Error(`The process measuring ${dist} ended with status ${child.status}`);
  }
  return JSON.parse(child.stdout) as Measurement[];
}

function summarize(samples: readonly number[]): Timing {
  const sorted = [...samples].sort((a, b) => a - b);
  return { minMs: sorted[0]!, p50Ms: sorted[Math.floor(sorted.length / 2)]! };
}

const iterations = Number(argument("--iterations") ?? 4);
const measured = argument("--measure");
if (measured !== undefined) {
  // A measuring process: its whole output is the measurements, for the process that started it.
  console.log(JSON.stringify(await measure(measured, iterations)));
} else {
  const headDist = argument("--dist") ?? fileURLToPath(new URL("../dist/", import.meta.url));
  const baseDist = argument("--base");
  const rounds = Number(argument("--rounds") ?? 2);

  // One process per build and round, the order swapped every round so that neither build always
  // runs on the warmer or the busier machine.
  const headSamples = new Map<string, number[]>();
  const baseSamples = new Map<string, number[]>();
  const unwritable = new Set<string>();
  for (let round = 0; round < rounds; round++) {
    const builds: (readonly [dist: string, isHead: boolean])[] =
      baseDist === undefined
        ? [[headDist, true]]
        : [
            [headDist, true],
            [baseDist, false],
          ];
    if (round % 2 === 1) builds.reverse();
    for (const [dist, isHead] of builds) {
      for (const { suite, variant, samples } of measureInProcess(dist, iterations)) {
        const key = `${suite}\n${variant}`;
        if (samples === null) {
          // A base from before streaming or auto widths has no such row; the build under test
          // must write every variant.
          if (isHead) throw new Error(`The build under test could not write ${suite}, ${variant}`);
          unwritable.add(key);
          continue;
        }
        const into = isHead ? headSamples : baseSamples;
        into.set(key, [...(into.get(key) ?? []), ...samples]);
      }
    }
  }

  const benchmarks: BenchmarkResult[] = [];
  for (const [suite] of suites) {
    for (const variant of VARIANTS) {
      const key = `${suite}\n${variant}`;
      const base = unwritable.has(key) ? undefined : baseSamples.get(key);
      benchmarks.push({
        suite,
        variant,
        samples: rounds * iterations,
        head: summarize(headSamples.get(key)!),
        ...(base === undefined ? {} : { base: summarize(base) }),
      });
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
    `${run.runtime} ${run.version} on ${run.arch}, ${rounds} processes of ${iterations} runs per benchmark${baseDist === undefined ? "" : ", in turns with the base build"}`,
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
}
