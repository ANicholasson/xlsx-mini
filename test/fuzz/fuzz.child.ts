import { readFileSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { isDeepStrictEqual } from "node:util";
import { type Outcome, outcomeOf, render, send } from "../helpers/outcome.ts";
import { type Depth, mutate, randomOf, sheetHoldsValues } from "./mutate.ts";

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
      /** Sheets that had rows and read with none, whose part holds no value any more. */
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

const rowsOf = (outcome: Outcome, name: string) =>
  "sheets" in outcome ? outcome.sheets.find((sheet) => sheet.name === name)?.rows : undefined;

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
    const outcome = await outcomeOf(mutation.bytes);
    const refusal =
      "refused" in outcome
        ? outcome.refused
        : outcome.sheets.find((sheet) => sheet.refused !== undefined)?.refused;
    ended =
      refusal?.code ??
      (isDeepStrictEqual(outcome, original.outcome) ? "read" : "read other values");
    if (ended === "read other values" && mutation.same) {
      failures.push(`read values the original does not hold: ${render(outcome)}`);
    }
    for (const sheet of "sheets" in outcome ? outcome.sheets : []) {
      if (sheet.rows?.length !== 0 || !rowsOf(original.outcome, sheet.name)?.length) continue;
      const holds = sheetHoldsValues(mutation.bytes, sheet.name);
      if (holds === false) excused++;
      else {
        failures.push(
          `sheet ${JSON.stringify(sheet.name)} had rows and reads with none, ${holds ? "though its part holds a value" : "and its part cannot be followed"}`,
        );
      }
    }
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
