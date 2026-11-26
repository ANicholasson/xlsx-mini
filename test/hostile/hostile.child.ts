import { isDeepStrictEqual } from "node:util";
import { getHeapStatistics } from "node:v8";
import { outcomeOf, render, send } from "../helpers/outcome.ts";
import { hostileCases } from "./cases.ts";

/** What the child sends: `built` once the file is made, then one of the others. */
export type HostileMessage =
  | { readonly kind: "built"; readonly fileBytes: number; readonly heapLimit: number }
  | {
      readonly kind: "read";
      readonly same: boolean;
      readonly outcome: string;
      readonly expected: string;
      readonly milliseconds: number;
    }
  | { readonly kind: "failed"; readonly failure: string };

const fail = (failure: string) =>
  send({ kind: "failed", failure } satisfies HostileMessage).finally(() => process.exit(1));

process.on("unhandledRejection", (reason) => {
  void fail(
    `a rejection that nothing awaits: ${reason instanceof Error ? reason.stack : String(reason)}`,
  );
});

const name = process.argv[2];
const hostile = hostileCases.find((each) => each.name === name);
if (hostile === undefined) throw new Error(`no hostile case named ${name}`);

const file = hostile.file();
const expected = typeof hostile.outcome === "function" ? hostile.outcome() : hostile.outcome;
await send({
  kind: "built",
  fileBytes: file instanceof Blob ? file.size : file.length,
  heapLimit: getHeapStatistics().heap_size_limit,
} satisfies HostileMessage);
const started = performance.now();
try {
  const outcome = await outcomeOf(file, hostile.plan, hostile.limits);
  await send({
    kind: "read",
    same: isDeepStrictEqual(outcome, expected),
    outcome: render(outcome),
    expected: render(expected),
    milliseconds: performance.now() - started,
  } satisfies HostileMessage);
  process.disconnect?.();
} catch (error) {
  await fail(`not an XlsxReadError: ${error instanceof Error ? error.stack : String(error)}`);
}
