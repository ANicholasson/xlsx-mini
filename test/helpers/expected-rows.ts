import type { Cell, Coerce, Sheet } from "../../src/index.js";

type Expected = string | number | boolean | Date | null;

export type ExpectedRow = readonly [number: number, ...cells: Expected[]];

/** XML 1.0 section 2.2, the production `Char`: what a part can carry. */
function carried(text: string): string | null {
  let kept = "";
  for (const character of text) {
    const point = character.codePointAt(0) ?? 0;
    const allowed =
      point === 9 ||
      point === 10 ||
      point === 13 ||
      (point >= 32 && point < 0xd800) ||
      (point > 0xdfff && point < 0xfffe) ||
      point > 0xffff;
    if (allowed) kept += character;
  }
  return kept === "" ? null : kept;
}

interface Civil {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour?: number;
  readonly minute?: number;
  readonly second?: number;
  readonly millisecond?: number;
}

function valueOf(value: unknown, coerce: Coerce | undefined): Expected {
  if (value === null || value === undefined) return null;
  if (typeof value === "string") return carried(value);
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (coerce !== undefined) return valueOf(coerce(value as Date, "expected"), undefined);
  if (value instanceof Date) return new Date(value.getTime());
  const tag = (value as { readonly [Symbol.toStringTag]?: unknown })[Symbol.toStringTag];
  if (tag === "Temporal.Instant") {
    return new Date((value as { readonly epochMilliseconds: number }).epochMilliseconds);
  }
  if (typeof tag === "string" && tag.startsWith("Temporal.")) {
    const civil = value as Civil;
    const { year, month, day, hour = 0, minute = 0, second = 0, millisecond = 0 } = civil;
    return new Date(Date.UTC(year, month - 1, day, hour, minute, second, millisecond));
  }
  throw new Error(`no expected value for a ${typeof value}`);
}

function cellOf(cell: Cell, coerce: Coerce | undefined): Expected {
  if (typeof cell !== "object" || cell === null || cell instanceof Date) {
    return valueOf(cell, coerce);
  }
  if ("image" in cell) return null;
  if ("value" in cell) return valueOf(cell.value, coerce);
  if ("formula" in cell) return null;
  return valueOf(cell, coerce);
}

/**
 * The rows the reader gives for a sheet the writer was given, derived from the input alone. A
 * date-like value is the `Date` whose UTC fields are its wall clock to the millisecond: the writer
 * gives it a date format, and no fixture gives it a number format that shows no date.
 */
export function expectedRows(sheet: Sheet, coerce?: Coerce): ExpectedRow[] {
  const rows: ExpectedRow[] = [];
  let number = 0;
  for (const row of sheet.rows) {
    number++;
    const cells = ("cells" in row ? row.cells : row).map((cell) => cellOf(cell, coerce));
    while (cells.length > 0 && cells.at(-1) === null) cells.pop();
    if (cells.length > 0) rows.push([number, ...cells]);
  }
  return rows;
}
