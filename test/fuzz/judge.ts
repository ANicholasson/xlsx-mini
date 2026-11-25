import { isDeepStrictEqual } from "node:util";
import { type Outcome, render } from "../helpers/outcome.ts";
import { type Mutation, sheetHoldsValues, sheetListed } from "./mutate.ts";

export interface Judgment {
  /** `read`, `read other values`, or the code of the first refusal. */
  readonly ended: string;
  /** Sheets that had rows and read with none, or are not read, because the mutation took them away. */
  readonly excused: number;
  readonly failures: readonly string[];
}

/**
 * What a mutated package's read ended in, and what is wrong with it. A read fails when it holds
 * values the original does not though the mutation changed no part's content, and when a sheet that
 * had rows reads with none or is gone from the workbook, unless the package itself no longer holds
 * the sheet's values or no longer lists the sheet: a damaged workbook never reads as an empty one.
 */
export function judge(
  original: Outcome,
  outcome: Outcome,
  mutation: Pick<Mutation, "bytes" | "same">,
): Judgment {
  const refusal =
    "refused" in outcome
      ? outcome.refused
      : outcome.sheets.find((sheet) => sheet.refused !== undefined)?.refused;
  const ended =
    refusal?.code ?? (isDeepStrictEqual(outcome, original) ? "read" : "read other values");
  const failures: string[] = [];
  let excused = 0;
  if (ended === "read other values" && mutation.same) {
    failures.push(`read values the original does not hold: ${render(outcome)}`);
  }
  if ("sheets" in outcome && "sheets" in original) {
    for (const { name, rows } of original.sheets) {
      if (!rows?.length) continue;
      const sheet = outcome.sheets.find((each) => each.name === name);
      if (sheet === undefined) {
        const listed = sheetListed(mutation.bytes, name);
        if (listed === false) excused++;
        else {
          failures.push(
            `sheet ${JSON.stringify(name)} had rows and is not read, ${listed ? "though the workbook part lists it" : "and the workbook part cannot be followed"}`,
          );
        }
      } else if (sheet.rows?.length === 0) {
        const holds = sheetHoldsValues(mutation.bytes, name);
        if (holds === false) excused++;
        else {
          failures.push(
            `sheet ${JSON.stringify(name)} had rows and reads with none, ${holds ? "though its part holds a value" : "and its part cannot be followed"}`,
          );
        }
      }
    }
  }
  return { ended, excused, failures };
}
