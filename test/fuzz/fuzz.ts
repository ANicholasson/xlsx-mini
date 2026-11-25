// node test/fuzz/fuzz.ts [--seed N] [--mutations N | --mutation I] [--minutes M] [--issue FILE] [--run URL]
//
// Mutates the committed corpus from a seed and reads each mutation in a process of its own under a
// heap cap and a time bound per mutation. Prints what the mutations ended in, by depth, and a log
// line for each failure with the command that reads that mutation again. With --issue it writes the
// text of an issue about the failures to FILE. Exits with 1 when a mutation failed.
import { randomInt } from "node:crypto";
import { writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { bundled, supervise } from "../helpers/child.ts";
import type { FuzzMessage } from "./fuzz.child.ts";
import type { Depth } from "./mutate.ts";

const here = dirname(fileURLToPath(import.meta.url));

/**
 * The old space of the process that reads the mutations: four times the 16 MiB that 20 000 of them
 * needed on Node 24 and 26. It fails a reader that sizes the shared-string table from its `count`,
 * which 10 of the first 20 000 mutations of seed 91 set past four billion, the first at 1 148.
 */
const HEAP_MEBIBYTES = 64;

/**
 * The time one mutation may take: more than a hundred times the slowest of 45 000 on the machine
 * that set it, 16 ms. It fails a reader whose walk of a damaged directory stops moving forward.
 */
const MUTATION_MILLISECONDS = 2_000;

/** A run stops after this many failures, so one that hangs on every mutation still ends soon. */
const MOST_FAILURES = 20;

/**
 * The run in the checks: a fixed seed, so a failure there is a change in the code, not a draw. Of
 * the seeds 90 to 98, the first whose 5 000 mutations empty a sheet by changing its text, so the
 * run reaches the exclusion it states.
 */
export const CHECKED_RUN = { seed: 91, count: 5_000 } as const;

export interface FuzzRun {
  readonly seed: number;
  readonly first?: number;
  readonly count: number;
  /** The time, in milliseconds since the epoch, after which no mutation starts. */
  readonly until?: number;
}

/**
 * One mutation that failed the run: it ended in neither a refusal with a code nor a read, left a
 * rejection that nothing awaits, read a sheet that holds a value as one without rows, or read
 * damaged bytes as other values.
 */
export interface FuzzFailure {
  readonly seed: number;
  readonly index: number;
  readonly workbook: string;
  readonly depth: Depth;
  readonly mutation: string;
  readonly failure: string;
}

/** What the mutations of one depth ended in. */
interface Tally {
  mutations: number;
  /** Read with the original's values. */
  read: number;
  /** Read with other values: a package that states another workbook. */
  readOther: number;
  refused: Record<string, number>;
  /** Ended in neither a refusal with a code nor a read. */
  failed: number;
  /** Sheets that had rows and read with none or are not read, because the mutation took them away. */
  excused: number;
}

export interface FuzzReport {
  readonly seed: number;
  readonly first: number;
  readonly milliseconds: number;
  readonly tallies: Readonly<Record<Depth, Tally>>;
  /** What each mutation ended in, by its index. */
  readonly ended: ReadonlyMap<number, string>;
  readonly failures: readonly FuzzFailure[];
  /** The mutation that took longest, and its time in milliseconds. */
  readonly slowest: { readonly index: number; readonly milliseconds: number };
}

const tally = (): Tally => ({
  mutations: 0,
  read: 0,
  readOther: 0,
  refused: {},
  failed: 0,
  excused: 0,
});

/**
 * Reads the mutations `first` to `first + count` of `seed`, or those that start before `until`. A
 * process that ends while it reads one, or passes the time bound, fails that mutation; a new process
 * goes on from the next.
 */
export async function fuzz(run: FuzzRun): Promise<FuzzReport> {
  const started = performance.now();
  const first = run.first ?? 0;
  const end = first + run.count;
  const tallies = { bytes: tally(), records: tally(), text: tally() };
  const ended = new Map<number, string>();
  const failures: FuzzFailure[] = [];
  let slowest = { index: first, milliseconds: 0 };
  const child = await bundled(join(here, "fuzz.child.ts"));
  try {
    for (
      let next = first;
      next < end && Date.now() < (run.until ?? Infinity) && failures.length < MOST_FAILURES;
    ) {
      const { messages, exit, overran, stderr } = await supervise<FuzzMessage>(
        child.script,
        [
          join(here, "..", "corpus"),
          String(run.seed),
          String(next),
          String(end - next),
          String(run.until ?? Number.MAX_SAFE_INTEGER),
          String(MOST_FAILURES - failures.length),
        ],
        HEAP_MEBIBYTES,
        ({ kind }) => (kind === "started" ? MUTATION_MILLISECONDS : undefined),
      );
      for (const message of messages) {
        if (message.kind !== "done") continue;
        const counts = tallies[message.depth];
        counts.mutations++;
        counts.excused += message.excused;
        if (message.ended === "read") counts.read++;
        else if (message.ended === "read other values") counts.readOther++;
        else if (message.ended === "failed") counts.failed++;
        else counts.refused[message.ended] = (counts.refused[message.ended] ?? 0) + 1;
        ended.set(message.index, message.ended);
        if (message.milliseconds > slowest.milliseconds) {
          slowest = { index: message.index, milliseconds: message.milliseconds };
        }
        if (message.failure !== undefined) {
          const { index, workbook, depth, mutation, failure } = message;
          failures.push({ seed: run.seed, index, workbook, depth, mutation, failure });
        }
        next = message.index + 1;
      }
      const last = messages.at(-1);
      if (last?.kind === "started") {
        failures.push({
          seed: run.seed,
          index: last.index,
          workbook: last.workbook,
          depth: last.depth,
          mutation: last.mutation,
          failure:
            overran === undefined
              ? `the process ended (${exit}) while it read the mutation: ${stderr}`
              : `passed its time bound of ${MUTATION_MILLISECONDS} ms`,
        });
        ended.set(last.index, "failed");
        tallies[last.depth].mutations++;
        tallies[last.depth].failed++;
        next = last.index + 1;
      } else if (exit !== "code 0") {
        throw new Error(`the fuzz process ended (${exit}) between mutations: ${stderr}`);
      } else if (messages.length === 0 && Date.now() < (run.until ?? Infinity)) {
        throw new Error(
          `the fuzz process ended without reading mutation ${next} of seed ${run.seed}`,
        );
      }
    }
  } finally {
    child.remove();
  }
  return {
    seed: run.seed,
    first,
    milliseconds: performance.now() - started,
    tallies,
    ended,
    failures,
    slowest,
  };
}

const rerunOf = ({ seed, index }: FuzzFailure) =>
  `node test/fuzz/fuzz.ts --seed ${seed} --mutation ${index}`;

/** One failure as the log prints it: everything needed to read it again. */
export function logLine(failure: FuzzFailure): string {
  const { seed, index, workbook, depth, mutation } = failure;
  return `seed ${seed} mutation ${index} of ${workbook} (${depth}: ${mutation}): ${failure.failure}\n  ${rerunOf(failure)}`;
}

function summaryOf(report: FuzzReport): string {
  const lines = Object.entries(report.tallies).map(([depth, counts]) => {
    const refused = Object.entries(counts.refused)
      .sort(([, a], [, b]) => b - a)
      .map(([code, number]) => `${code} ${number}`)
      .join(", ");
    return `${depth}: ${counts.mutations} mutations, ${counts.read} read as the original, ${counts.readOther} read as another workbook, ${counts.excused} sheets emptied by the mutation, ${counts.failed} failed; refused ${refused || "none"}`;
  });
  const seconds = (report.milliseconds / 1000).toFixed(1);
  const slowest = `the slowest, mutation ${report.slowest.index}, in ${report.slowest.milliseconds.toFixed(0)} ms`;
  return [
    `seed ${report.seed}, from mutation ${report.first}, in ${seconds} s; ${slowest}${report.failures.length >= MOST_FAILURES ? `; stopped at ${MOST_FAILURES} failures` : ""}`,
    ...lines,
  ].join("\n");
}

/** Marks the issue the scheduled run opens, so a later failure finds it and comments on it. */
export const ISSUE_MARKER = "<!-- xlsx-mini-fuzz -->";

const LISTED_FAILURES = 50;

/** The text of an issue, or of a comment on one, about a run's failures. */
export function issueOf(report: FuzzReport, runUrl?: string): string {
  const failures = report.failures.slice(0, LISTED_FAILURES);
  const more = report.failures.length - failures.length;
  return [
    ISSUE_MARKER,
    `The scheduled fuzz run${runUrl === undefined ? "" : ` (${runUrl})`} read mutations of the committed corpus from seed ${report.seed}. ${report.failures.length} of them ended in neither a refusal with a code nor a clean read.`,
    "",
    "Each failure, and the command that reads it again:",
    "",
    "```text",
    ...failures.map(logLine),
    ...(more > 0 ? [`and ${more} more`] : []),
    "```",
    "",
    "What the run read:",
    "",
    "```text",
    summaryOf(report),
    "```",
    "",
  ].join("\n");
}

function option(name: string): string | undefined {
  const at = process.argv.indexOf(`--${name}`);
  return at < 0 ? undefined : process.argv[at + 1];
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const seed = Number(option("seed") ?? randomInt(2 ** 32));
  if (!Number.isSafeInteger(seed) || seed < 0) throw new Error("--seed takes a whole number");
  const single = option("mutation");
  const minutes = option("minutes");
  const report = await fuzz({
    seed,
    first: single === undefined ? 0 : Number(single),
    count: single === undefined ? Number(option("mutations") ?? CHECKED_RUN.count) : 1,
    ...(minutes === undefined ? {} : { until: Date.now() + Number(minutes) * 60_000 }),
  });
  console.log(summaryOf(report));
  for (const failure of report.failures) console.error(logLine(failure));
  const issue = option("issue");
  if (issue !== undefined && report.failures.length > 0) {
    writeFileSync(issue, issueOf(report, option("run")));
  }
  if (report.failures.length > 0) process.exitCode = 1;
}
