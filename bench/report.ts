// Turns benchmark results into the pull-request comment: one table per runtime, the pull request
// against the base build timed alongside it. Pure functions first; the command line at the bottom.
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

export type Variant = "deflate" | "store";

/** The minimum and the median of a benchmark's runs, in milliseconds. */
export interface Timing {
  readonly minMs: number;
  readonly p50Ms: number;
}

/** One benchmark of one run: a suite written one way, for the build under test and its base. */
export interface BenchmarkResult {
  readonly suite: string;
  readonly variant: Variant;
  readonly samples: number;
  readonly head: Timing;
  readonly base?: Timing;
}

/** Every benchmark of one runtime. */
export interface RunResult {
  readonly runtime: string;
  readonly version: string;
  readonly arch: string;
  readonly commit: string;
  readonly baseCommit?: string;
  readonly date: string;
  readonly benchmarks: readonly BenchmarkResult[];
}

/** Changes inside this band are noise, even with both builds on one machine. */
export const NOISE_BAND = 0.05;
/** A change beyond this band is called a regression or an improvement outright. */
export const ALERT_BAND = 0.15;

/** The marker that lets the workflow find and update its own comment. */
export const COMMENT_MARKER = "<!-- xlsx-mini-benchmarks -->";

const RUNTIME_ORDER = ["node", "bun", "deno"];
const RUNTIME_NAMES: Readonly<Record<string, string>> = { node: "Node", bun: "Bun", deno: "Deno" };

function runtimeRank(runtime: string): number {
  const index = RUNTIME_ORDER.indexOf(runtime);
  return index === -1 ? RUNTIME_ORDER.length : index;
}

function runtimeTitle(run: RunResult): string {
  return `${RUNTIME_NAMES[run.runtime] ?? run.runtime} ${run.version}`;
}

function ms(value: number): string {
  return `${value.toFixed(1)} ms`;
}

function shortCommit(commit: string | undefined): string {
  return commit === undefined ? "" : `\`${commit.slice(0, 7)}\``;
}

/** The change column: a signed percentage, flagged outside the noise band. */
export function describeChange(baseMs: number, headMs: number): string {
  const ratio = headMs / baseMs - 1;
  const percent = `${ratio >= 0 ? "+" : ""}${(ratio * 100).toFixed(1)} %`;
  if (ratio > ALERT_BAND) return `🔴 ${percent}`;
  if (ratio > NOISE_BAND) return `⚠️ ${percent}`;
  if (ratio < -NOISE_BAND) return `🟢 ${percent}`;
  return percent;
}

interface ReportOptions {
  readonly baseLabel: string;
  readonly headLabel: string;
}

/** The pull-request comment: a table per runtime, minimum of the runs, both builds side by side. */
export function renderReport(runs: readonly RunResult[], options: ReportOptions): string {
  const ordered = [...runs].sort((a, b) => runtimeRank(a.runtime) - runtimeRank(b.runtime));
  const first = ordered[0];
  const samples = first?.benchmarks[0]?.samples ?? 0;
  const lines: string[] = [
    COMMENT_MARKER,
    "## Benchmarks",
    "",
    `Write time per suite: ${options.headLabel} and \`${options.baseLabel}\` built and timed on the same runner, call by call in turns, the minimum of ${samples} runs each. ` +
      `Changes inside ±${NOISE_BAND * 100} % are noise; ⚠️ marks more than ${NOISE_BAND * 100} % slower, 🔴 more than ${ALERT_BAND * 100} % slower, 🟢 more than ${NOISE_BAND * 100} % faster.`,
    "",
  ];
  for (const run of ordered) {
    lines.push(
      `### ${runtimeTitle(run)}`,
      "",
      `| Suite | Variant | ${options.baseLabel} | ${options.headLabel} | Change |`,
      "| --- | --- | ---: | ---: | ---: |",
    );
    for (const b of run.benchmarks) {
      const base = b.base === undefined ? "—" : ms(b.base.minMs);
      const change = b.base === undefined ? "—" : describeChange(b.base.minMs, b.head.minMs);
      lines.push(`| ${b.suite} | ${b.variant} | ${base} | ${ms(b.head.minMs)} | ${change} |`);
    }
    lines.push("");
  }
  lines.push(
    `<sub>${options.headLabel} ${shortCommit(first?.commit)}` +
      (first?.baseCommit === undefined
        ? ""
        : ` against ${options.baseLabel} ${shortCommit(first.baseCommit)}`) +
      `; ${samples} runs each, minimum.</sub>`,
  );
  return `${lines.join("\n")}\n`;
}

/** Reads every run file in a directory; an absent directory is no runs. */
export function readRuns(dir: string): RunResult[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((name) => name.endsWith(".json"))
    .map((name) => JSON.parse(readFileSync(join(dir, name), "utf8")) as Partial<RunResult>)
    .filter((run): run is RunResult => Array.isArray(run.benchmarks));
}

function argument(argv: readonly string[], name: string): string | undefined {
  const index = argv.indexOf(name);
  return index === -1 ? undefined : argv[index + 1];
}

/** `--head <dir>` with `--out <file>` renders the comment from the run files in the directory. */
export function main(argv: readonly string[]): void {
  const headDir = argument(argv, "--head");
  const out = argument(argv, "--out");
  if (headDir === undefined || out === undefined) {
    throw new Error("--head <dir> and --out <file> are required");
  }
  const report = renderReport(readRuns(headDir), {
    baseLabel: argument(argv, "--base-label") ?? "main",
    headLabel: argument(argv, "--head-label") ?? "this pull request",
  });
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, report);
}

if ((import.meta as { main?: boolean }).main === true) main(process.argv.slice(2));
