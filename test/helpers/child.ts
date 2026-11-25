import { execFile, fork } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";
import { rolldown } from "rolldown";

/** A child bundled into one script, in a folder of its own that `remove` deletes. */
export interface Bundled {
  readonly script: string;
  readonly remove: () => void;
}

/** Bundles `child`, a TypeScript file, into one script: Node 20 does not run TypeScript. */
export async function bundled(child: string): Promise<Bundled> {
  const directory = mkdtempSync(join(tmpdir(), "xlsx-mini-child-"));
  const build = await rolldown({ input: child, platform: "node", logLevel: "silent" });
  const { output } = await build.generate({ format: "esm" });
  await build.close();
  const script = join(directory, "child.mjs");
  writeFileSync(script, output[0].code);
  return { script, remove: () => rmSync(directory, { recursive: true, force: true }) };
}

/**
 * Runs `child`, a TypeScript file, in a process of its own whose heap is held to `heapMebibytes`
 * and which has a collector to call, and parses the line of JSON it prints. The process fails, and
 * so does the promise, when the heap does not hold what the child reads.
 */
export async function reportOf(child: string, heapMebibytes: number): Promise<unknown> {
  const { script, remove } = await bundled(child);
  try {
    const { stdout } = await promisify(execFile)(
      process.execPath,
      [`--max-old-space-size=${heapMebibytes}`, "--expose-gc", script],
      { encoding: "utf8" },
    );
    return JSON.parse(stdout);
  } finally {
    remove();
  }
}

/** How a supervised child ended, with every message it sent. */
export interface Supervised<Message> {
  readonly messages: readonly Message[];
  /** `code 0` for a child that finished, otherwise its exit code or the signal that ended it. */
  readonly exit: string;
  /** The message whose time bound passed, after which the child was killed. */
  readonly overran: Message | undefined;
  /**
   * What the child wrote to its standard error, cut short: the heap's `FATAL ERROR` when it ran out,
   * otherwise its first 500 characters.
   */
  readonly stderr: string;
}

/**
 * Runs a bundled script in a process of its own whose heap is held to `heapMebibytes`, and collects
 * the messages it sends, which may hold any value structured cloning carries, a `Date` and `NaN`
 * included. Each message may start a time bound: `bound` gives it in milliseconds, or `undefined`
 * for none, and the next message ends it. A child still running when its bound passes is killed.
 */
export function supervise<Message>(
  script: string,
  args: readonly string[],
  heapMebibytes: number,
  bound: (message: Message) => number | undefined,
): Promise<Supervised<Message>> {
  return new Promise((resolve, reject) => {
    const child = fork(script, args, {
      execArgv: [`--max-old-space-size=${heapMebibytes}`],
      serialization: "advanced",
      stdio: ["ignore", "ignore", "pipe", "ipc"],
    });
    const messages: Message[] = [];
    let overran: Message | undefined;
    let timer: ReturnType<typeof setTimeout> | undefined;
    let stderr = "";
    child.stderr?.setEncoding("utf8");
    let fatal: string | undefined;
    let before = "";
    child.stderr?.on("data", (text: string) => {
      fatal ??= /FATAL ERROR.*/.exec(before + text)?.[0];
      before = text.slice(-200);
      if (stderr.length < 8_000) stderr += text;
    });
    child.on("message", (message: Message) => {
      messages.push(message);
      clearTimeout(timer);
      const milliseconds = bound(message);
      if (milliseconds === undefined) return;
      timer = setTimeout(() => {
        overran = message;
        child.kill("SIGKILL");
      }, milliseconds);
    });
    child.on("error", reject);
    child.on("close", (code, signal) => {
      clearTimeout(timer);
      resolve({
        messages,
        exit: signal ?? `code ${code}`,
        overran,
        stderr: fatal ?? stderr.slice(0, 500),
      });
    });
  });
}
