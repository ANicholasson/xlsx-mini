// Turns benchmark results into the pull-request comment, the history and the README chart.
// Pure functions first; the command line at the bottom wires files to them.
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";

export type Variant = "deflate" | "store";

/** One benchmark of one run: a suite written one way, timed in milliseconds. */
export interface BenchmarkResult {
  readonly suite: string;
  readonly variant: Variant;
  readonly samples: number;
  readonly p50Ms: number;
  readonly minMs: number;
  readonly meanMs: number;
}

/** Every benchmark of one runtime on one commit. */
export interface RunResult {
  readonly runtime: string;
  readonly version: string;
  readonly arch: string;
  readonly commit: string;
  readonly date: string;
  readonly benchmarks: readonly BenchmarkResult[];
}

/** One point of history: the median of one benchmark on one commit of `main`. */
export interface HistoryEntry {
  readonly commit: string;
  readonly date: string;
  readonly runtime: string;
  readonly suite: string;
  readonly variant: Variant;
  readonly p50Ms: number;
}

/** Changes inside this band are noise on a shared runner. */
export const NOISE_BAND = 0.1;
/** A change beyond this band is called a regression or an improvement outright. */
export const ALERT_BAND = 0.25;
/** How many commits of history the chart shows. */
export const HISTORY_POINTS = 20;
/** How many entries the history keeps per benchmark. */
export const HISTORY_LIMIT = 100;

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

/** The change column: a signed percentage, flagged outside the noise band. */
export function describeChange(baseMs: number, headMs: number): string {
  const ratio = headMs / baseMs - 1;
  const percent = `${ratio >= 0 ? "+" : ""}${(ratio * 100).toFixed(1)} %`;
  if (ratio > ALERT_BAND) return `🔴 ${percent}`;
  if (ratio > NOISE_BAND) return `⚠️ ${percent}`;
  if (ratio < -NOISE_BAND) return `🟢 ${percent}`;
  return percent;
}

function findBenchmark(run: RunResult | undefined, suite: string, variant: Variant) {
  return run?.benchmarks.find((b) => b.suite === suite && b.variant === variant);
}

function chartCeiling(values: readonly number[]): number {
  const top = Math.max(...values, 1) * 1.1;
  const step = top > 1000 ? 500 : top > 200 ? 100 : top > 50 ? 25 : 5;
  return Math.ceil(top / step) * step;
}

function quoted(labels: readonly string[]): string {
  return `[${labels.map((label) => `"${label.replaceAll('"', "'")}"`).join(", ")}]`;
}

function numbers(values: readonly number[]): string {
  return `[${values.map((value) => value.toFixed(1)).join(", ")}]`;
}

interface ReportOptions {
  readonly baseLabel: string;
  readonly headLabel: string;
}

/**
 * The pull-request comment: a table per runtime with the change against the base, a chart of the
 * base against the head for the first runtime, and a history chart when history exists.
 */
export function renderReport(
  head: readonly RunResult[],
  base: readonly RunResult[],
  history: readonly HistoryEntry[],
  options: ReportOptions,
): string {
  const runs = [...head].sort((a, b) => runtimeRank(a.runtime) - runtimeRank(b.runtime));
  const lines: string[] = [COMMENT_MARKER, "## Benchmarks", ""];
  lines.push(
    `Write time per suite, median of the runs on a shared GitHub runner, against \`${options.baseLabel}\`. ` +
      `Changes inside ±${NOISE_BAND * 100} % are noise; ⚠️ marks more than ${NOISE_BAND * 100} % slower, 🔴 more than ${ALERT_BAND * 100} % slower, 🟢 more than ${NOISE_BAND * 100} % faster.`,
    "",
  );
  if (base.length === 0) {
    lines.push(`No results on \`${options.baseLabel}\` yet; the next push there records them.`, "");
  }

  for (const run of runs) {
    const baseRun = base.find((b) => b.runtime === run.runtime);
    lines.push(`### ${runtimeTitle(run)}`, "");
    lines.push(
      `| Suite | Variant | ${options.baseLabel} | ${options.headLabel} | Change |`,
      "| --- | --- | ---: | ---: | ---: |",
    );
    for (const b of run.benchmarks) {
      const before = findBenchmark(baseRun, b.suite, b.variant);
      const baseCell = before === undefined ? "—" : ms(before.p50Ms);
      const change = before === undefined ? "—" : describeChange(before.p50Ms, b.p50Ms);
      lines.push(`| ${b.suite} | ${b.variant} | ${baseCell} | ${ms(b.p50Ms)} | ${change} |`);
    }
    lines.push("");
  }

  const first = runs[0];
  if (first !== undefined) {
    const baseRun = base.find((b) => b.runtime === first.runtime);
    for (const variant of ["deflate", "store"] as const) {
      const benchmarks = first.benchmarks.filter((b) => b.variant === variant);
      if (benchmarks.length === 0) continue;
      const headValues = benchmarks.map((b) => b.p50Ms);
      const baseValues = benchmarks.map(
        (b) => findBenchmark(baseRun, b.suite, b.variant)?.p50Ms ?? 0,
      );
      const hasBase = baseRun !== undefined && baseValues.every((value) => value > 0);
      lines.push(
        "```mermaid",
        "xychart-beta",
        `    title "${runtimeTitle(first)}, ${variant}: ${hasBase ? `${options.baseLabel} as bars, ${options.headLabel} as the line` : options.headLabel}, ms"`,
        `    x-axis ${quoted(benchmarks.map((b) => b.suite))}`,
        `    y-axis "ms" 0 --> ${chartCeiling([...headValues, ...baseValues])}`,
        `    bar ${numbers(hasBase ? baseValues : headValues)}`,
      );
      if (hasBase) lines.push(`    line ${numbers(headValues)}`);
      lines.push("```", "");
    }
  }

  const historyChart = renderHistoryChart(history, first?.runtime ?? RUNTIME_ORDER[0]!);
  if (historyChart !== undefined) lines.push(historyChart, "");

  lines.push(
    `<sub>${options.headLabel} ${shortCommit(first?.commit)}` +
      (base[0] === undefined
        ? ""
        : ` against ${options.baseLabel} ${shortCommit(base[0].commit)}`) +
      `; ${first?.benchmarks[0]?.samples ?? 0} runs per benchmark.</sub>`,
  );
  return `${lines.join("\n")}\n`;
}

function shortCommit(commit: string | undefined): string {
  return commit === undefined ? "" : `\`${commit.slice(0, 7)}\``;
}

/** The last commits of history on one runtime, the deflate variant of every suite as a line. */
export function renderHistoryChart(
  history: readonly HistoryEntry[],
  runtime: string,
): string | undefined {
  const entries = history.filter((e) => e.runtime === runtime && e.variant === "deflate");
  if (entries.length === 0) return undefined;
  const commits = [...new Set(entries.map((e) => e.commit))].slice(-HISTORY_POINTS);
  const suites = [...new Set(entries.map((e) => e.suite))];
  const series = suites.map((suite) =>
    commits.map(
      (commit) => entries.find((e) => e.commit === commit && e.suite === suite)?.p50Ms ?? 0,
    ),
  );
  const lines = [
    "```mermaid",
    "xychart-beta",
    `    title "History on main: ${RUNTIME_NAMES[runtime] ?? runtime}, deflate, ms"`,
    `    x-axis ${quoted(commits.map((c) => c.slice(0, 7)))}`,
    `    y-axis "ms" 0 --> ${chartCeiling(series.flat())}`,
  ];
  for (const values of series) lines.push(`    line ${numbers(values)}`);
  lines.push("```", "", `<sub>Lines in order: ${suites.join(", ")}.</sub>`);
  return lines.join("\n");
}

/** Adds the runs to history, one entry per benchmark, newest last, capped per benchmark. */
export function appendHistory(
  history: readonly HistoryEntry[],
  runs: readonly RunResult[],
): HistoryEntry[] {
  const added: HistoryEntry[] = runs.flatMap((run) =>
    run.benchmarks.map((b) => ({
      commit: run.commit,
      date: run.date,
      runtime: run.runtime,
      suite: b.suite,
      variant: b.variant,
      p50Ms: b.p50Ms,
    })),
  );
  const key = (e: HistoryEntry) => `${e.commit} ${e.runtime} ${e.suite} ${e.variant}`;
  const keys = new Set(added.map(key));
  const kept = history.filter((e) => !keys.has(key(e)));
  const all = [...kept, ...added].sort((a, b) => a.date.localeCompare(b.date));
  const perBenchmark = new Map<string, number>();
  const result: HistoryEntry[] = [];
  for (const entry of [...all].reverse()) {
    const id = `${entry.runtime} ${entry.suite} ${entry.variant}`;
    const count = perBenchmark.get(id) ?? 0;
    if (count >= HISTORY_LIMIT) continue;
    perBenchmark.set(id, count + 1);
    result.push(entry);
  }
  return result.reverse();
}

const SVG_COLORS = ["#2f81f7", "#3fb950", "#d29922"];

/** A grouped bar chart of the latest deflate medians per runtime, for the README. */
export function renderSvg(latest: readonly RunResult[]): string {
  const runs = [...latest].sort((a, b) => runtimeRank(a.runtime) - runtimeRank(b.runtime));
  const suites = [...new Set(runs.flatMap((r) => r.benchmarks.map((b) => b.suite)))];
  const values = runs.map((run) =>
    suites.map((suite) => findBenchmark(run, suite, "deflate")?.p50Ms ?? 0),
  );
  const top = chartCeiling(values.flat());
  const width = 720;
  const height = 300;
  const left = 56;
  const bottom = 250;
  const plotHeight = 200;
  const groupWidth = (width - left - 16) / Math.max(runs.length, 1);
  const barWidth = Math.min(48, (groupWidth - 24) / Math.max(suites.length, 1));
  const parts: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" font-family="ui-sans-serif, system-ui, sans-serif" font-size="12">`,
    `<rect width="${width}" height="${height}" fill="#ffffff"/>`,
    `<text x="${left}" y="22" font-size="14" font-weight="600" fill="#1f2328">Write time per suite, median, deflate, ms (lower is better)</text>`,
  ];
  for (let tick = 0; tick <= 4; tick++) {
    const value = (top * tick) / 4;
    const y = bottom - (plotHeight * tick) / 4;
    parts.push(
      `<line x1="${left}" y1="${y}" x2="${width - 16}" y2="${y}" stroke="#d0d7de"/>`,
      `<text x="${left - 6}" y="${y + 4}" text-anchor="end" fill="#57606a">${Math.round(value)}</text>`,
    );
  }
  runs.forEach((run, r) => {
    const x0 = left + r * groupWidth + 12;
    suites.forEach((suite, s) => {
      const value = values[r]![s]!;
      const barHeight = top === 0 ? 0 : (plotHeight * value) / top;
      const x = x0 + s * barWidth;
      parts.push(
        `<rect x="${x}" y="${bottom - barHeight}" width="${barWidth - 4}" height="${barHeight}" fill="${SVG_COLORS[s % SVG_COLORS.length]}"><title>${suite}: ${ms(value)}</title></rect>`,
      );
    });
    parts.push(
      `<text x="${x0 + (suites.length * barWidth) / 2}" y="${bottom + 18}" text-anchor="middle" fill="#1f2328">${runtimeTitle(run)}</text>`,
    );
  });
  suites.forEach((suite, s) => {
    const x = left + s * 160;
    parts.push(
      `<rect x="${x}" y="${height - 22}" width="12" height="12" fill="${SVG_COLORS[s % SVG_COLORS.length]}"/>`,
      `<text x="${x + 18}" y="${height - 12}" fill="#57606a">${suite}</text>`,
    );
  });
  parts.push("</svg>");
  return `${parts.join("\n")}\n`;
}

/** Reads every `*.json` run file in a directory; an absent directory is no runs. */
export function readRuns(dir: string): RunResult[] {
  if (!existsSync(dir)) return [];
  return readdirSync(dir)
    .filter((name) => name.endsWith(".json"))
    .map((name) => JSON.parse(readFileSync(join(dir, name), "utf8")) as Partial<RunResult>)
    .filter((run): run is RunResult => Array.isArray(run.benchmarks));
}

export function readHistory(file: string): HistoryEntry[] {
  return existsSync(file) ? (JSON.parse(readFileSync(file, "utf8")) as HistoryEntry[]) : [];
}

function argument(argv: readonly string[], name: string): string | undefined {
  const index = argv.indexOf(name);
  return index === -1 ? undefined : argv[index + 1];
}

function write(file: string, content: string): void {
  mkdirSync(dirname(file), { recursive: true });
  writeFileSync(file, content);
}

/**
 * `--head <dir>` with `--out <file>` renders the comment, against `--base <dir>` and
 * `--history <file>` when given. `--append-history <file>` adds the head runs to that file, and
 * `--svg <file>` draws the chart of the head runs.
 */
export function main(argv: readonly string[]): void {
  const headDir = argument(argv, "--head");
  if (headDir === undefined) throw new Error("--head <dir> is required");
  const head = readRuns(headDir);
  const historyFile = argument(argv, "--history") ?? argument(argv, "--append-history");
  let history = historyFile === undefined ? [] : readHistory(historyFile);

  const appendTo = argument(argv, "--append-history");
  if (appendTo !== undefined) {
    history = appendHistory(history, head);
    write(appendTo, `${JSON.stringify(history)}\n`);
  }
  const out = argument(argv, "--out");
  if (out !== undefined) {
    const baseDir = argument(argv, "--base");
    const base = baseDir === undefined ? [] : readRuns(baseDir);
    write(
      out,
      renderReport(head, base, history, {
        baseLabel: argument(argv, "--base-label") ?? "main",
        headLabel: argument(argv, "--head-label") ?? "this pull request",
      }),
    );
  }
  const svg = argument(argv, "--svg");
  if (svg !== undefined) write(svg, renderSvg(head));
}

if ((import.meta as { main?: boolean }).main === true) main(process.argv.slice(2));
