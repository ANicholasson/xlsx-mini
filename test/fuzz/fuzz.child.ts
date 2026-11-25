import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { type Outcome, outcomeOf, send } from "../helpers/outcome.ts";
import { judge } from "./judge.ts";
import { type Depth, mutate, randomOf } from "./mutate.ts";

interface Drawn {
  readonly index: number;
  readonly workbook: string;
  readonly depth: Depth;
  readonly mutation: string;
}

/** What the child sends: `started` before a mutation is read, `done` once it ended. */
export type FuzzMessage =
  | ({ readonly kind: "started" } & Drawn)
  | ({
      readonly kind: "done";
      /** `read`, `read other values`, the code of the first refusal, or `failed`. */
      readonly ended: string;
      /** Sheets that had rows and read with none or are not read, because the mutation took them away. */
      readonly excused: number;
      readonly failure: string | undefined;
      readonly milliseconds: number;
    } & Drawn);

const [corpus = "", ...numbers] = process.argv.slice(2);
const [seed, first, count, until, mostFailures] = numbers.map(Number);

const workbooks = readdirSync(corpus, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .flatMap(({ name }) =>
    readdirSync(join(corpus, name))
      .filter((file) => file.endsWith(".xlsx"))
      .map((file) => `${name}/${file}`),
  )
  .sort();

const originals = new Map<string, { readonly bytes: Uint8Array; readonly outcome: Outcome }>();
async function originalOf(workbook: string) {
  let original = originals.get(workbook);
  if (original === undefined) {
    const bytes = new Uint8Array(readFileSync(join(corpus, workbook)));
    original = { bytes, outcome: await outcomeOf(bytes) };
    originals.set(workbook, original);
  }
  return original;
}

let unhandled: string | undefined;
process.on("unhandledRejection", (reason) => {
  unhandled = `a rejection that nothing awaits: ${String(reason)}`;
});

let failed = 0;
for (
  let index = first!;
  index < first! + count! && Date.now() < until! && failed < mostFailures!;
  index++
) {
  const random = randomOf(seed!, index);
  const workbook = workbooks[Math.floor(random() * workbooks.length)]!;
  const original = await originalOf(workbook);
  const mutation = mutate(original.bytes, random);
  const { depth, description } = mutation;
  await send({
    kind: "started",
    index,
    workbook,
    depth,
    mutation: description,
  } satisfies FuzzMessage);
  const started = performance.now();
  let ended = "failed";
  let excused = 0;
  const failures: string[] = [];
  try {
    const judgment = judge(original.outcome, await outcomeOf(mutation.bytes), mutation);
    ended = judgment.ended;
    excused = judgment.excused;
    failures.push(...judgment.failures);
  } catch (error) {
    failures.push(`not an XlsxReadError: ${String(error)}`);
  }
  if (unhandled !== undefined) failures.push(unhandled);
  unhandled = undefined;
  if (failures.length > 0) failed++;
  await send({
    kind: "done",
    index,
    workbook,
    depth: mutation.depth,
    mutation: mutation.description,
    ended,
    excused,
    failure: failures.length === 0 ? undefined : failures.join("; "),
    milliseconds: performance.now() - started,
  } satisfies FuzzMessage);
}
process.disconnect?.();
