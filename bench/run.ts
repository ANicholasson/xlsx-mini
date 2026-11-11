// Times the three suites against the built package on whichever runtime runs this file: Node, Bun
// or Deno. Every measurement runs in a process of its own, started by this one: the process loads
// one build, builds one suite, writes it one way once to warm up and then times that write. With
// `--base <dist>` the processes of the two builds take turns, one benchmark at a time, so both see
// the same machine within seconds of each other, and the report compares the minima. Two builds in
// one process share the engine's compiled code and its heap, and on the runner that made one of
// them write a streamed sheet at less than half its speed; apart, a build meets what a program
// using it meets, and a benchmark is not shaped by the ones before it.
// `pnpm bench` prints a table; `--json <file>` also writes the results, and `--samples <file>`
// every process's timings, for studying the harness itself.
import { spawnSync } from "node:child_process";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import type * as Api from "../src/index.ts";
import type { BenchmarkResult, RunResult, Timing, Variant } from "./report.ts";
import { type SuiteName, suites, withAutoWidths } from "./workbooks.ts";

type Write = () => Promise<unknown>;

const VARIANTS = ["deflate", "store", "stream", "auto"] as const satisfies readonly Variant[];

/** What one process measured: its timings, or null when the build cannot write the variant. */
type Samples = readonly number[] | null;

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
 * resolves it, builds the one suite, writes it once to warm up, then `iterations` timed times
 * with a collection before each where the host exposes one.
 */
async function measure(
  dist: string,
  suite: string,
  variant: Variant,
  iterations: number,
): Promise<Samples> {
  const api = (await import(pathToFileURL(resolve(dist, "index.js")).href)) as typeof Api;
  const workbook = suites.find(([name]) => name === suite)![1]();
  const written = variant === "auto" ? withAutoWidths(workbook) : workbook;
  const write: Write =
    variant === "deflate"
      ? () => api.writeXlsx(written)
      : variant === "stream"
        ? () => api.streamXlsx(written).pipeTo(new WritableStream())
        : () => api.writeXlsx(written, { compression: "none" });
  if (!(await succeeds(write))) return null;
  const samples: number[] = [];
  for (let i = 0; i < iterations; i++) {
    collectGarbage();
    samples.push(await time(write));
  }
  return samples;
}

/** Runs this file again on the same runtime, with the same engine flags, for one measurement. */
function measureInProcess(
  dist: string,
  suite: SuiteName,
  variant: Variant,
  iterations: number,
): Samples {
  const script = fileURLToPath(import.meta.url);
  const { runtime } = detectRuntime();
  // Every measuring process gets a collector to call, however this one was started; Bun has its own.
  const engineArguments =
    runtime === "deno"
      ? ["run", "--allow-all", "--no-lock", "--v8-flags=--expose-gc"]
      : runtime === "node"
        ? ["--expose-gc"]
        : [];
  const child = spawnSync(
    process.execPath,
    [
      ...engineArguments,
      script,
      "--measure",
      dist,
      "--suite",
      suite,
      "--variant",
      variant,
      "--iterations",
      String(iterations),
    ],
    { encoding: "utf8", stdio: ["ignore", "pipe", "inherit"] },
  );
  if (child.status !== 0) {
    throw new Error(`Measuring ${suite}, ${variant} of ${dist} ended with status ${child.status}`);
  }
  return JSON.parse(child.stdout) as Samples;
}

function summarize(samples: readonly number[]): Timing {
  const sorted = [...samples].sort((a, b) => a - b);
  return { minMs: sorted[0]!, p50Ms: sorted[Math.floor(sorted.length / 2)]! };
}

const iterations = Number(argument("--iterations") ?? 4);
const measured = argument("--measure");
if (measured !== undefined) {
  // A measuring process: its whole output is its timings, for the process that started it.
  const samples = await measure(
    measured,
    argument("--suite")!,
    argument("--variant") as Variant,
    iterations,
  );
  console.log(JSON.stringify(samples));
} else {
  const headDist = argument("--dist") ?? fileURLToPath(new URL("../dist/", import.meta.url));
  const baseDist = argument("--base");
  const rounds = Number(argument("--rounds") ?? 3);

  // Per round, one process per benchmark and build. The two builds run one after the other, and
  // which goes first swaps from one benchmark and one round to the next, so that neither always
  // runs on the warmer or the busier machine.
  const headSamples = new Map<string, number[]>();
  const baseSamples = new Map<string, number[]>();
  const unwritable = new Set<string>();
  const processes: object[] = [];
  let turn = 0;
  for (let round = 0; round < rounds; round++) {
    for (const [suite] of suites) {
      for (const variant of VARIANTS) {
        const key = `${suite}\n${variant}`;
        const builds: (readonly [dist: string, isHead: boolean])[] =
          baseDist === undefined || unwritable.has(key)
            ? [[headDist, true]]
            : [
                [headDist, true],
                [baseDist, false],
              ];
        if ((round + turn++) % 2 === 1) builds.reverse();
        for (const [dist, isHead] of builds) {
          const samples = measureInProcess(dist, suite, variant, iterations);
          processes.push({ build: isHead ? "head" : "base", round, suite, variant, samples });
          if (samples === null) {
            // A base from before streaming or auto widths has no such row; the build under test
            // must write every variant.
            if (isHead) throw new Error(`The build under test could not write ${key}`);
            unwritable.add(key);
            continue;
          }
          const into = isHead ? headSamples : baseSamples;
          into.set(key, [...(into.get(key) ?? []), ...samples]);
        }
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

  for (const [flag, content] of [
    ["--json", run],
    ["--samples", processes],
  ] as const) {
    const out = argument(flag);
    if (out === undefined) continue;
    mkdirSync(dirname(out), { recursive: true });
    writeFileSync(out, `${JSON.stringify(content, null, 2)}\n`);
    console.log(`written ${out}`);
  }
}
