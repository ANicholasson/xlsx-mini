import type { CellStyle } from "../../styles/index.js";

/**
 * Augment this interface to let your own types into cells, then convert them in `coerce`.
 * See the README section "Extending accepted value types".
 *
 * @example
 * declare module "xlsx-mini" {
 *   interface Register { cellValue: Decimal | bigint }
 * }
 */
// oxlint-disable-next-line typescript/no-empty-object-type, typescript/no-empty-interface
export interface Register {}

/** The value types a consumer registered, or `never` when none. */
export type Registered = Register extends { cellValue: infer T } ? T : never;

/** A Temporal.PlainDate, native or polyfilled, recognised by its tag; written as a date without time. */
export interface TemporalPlainDateLike {
  readonly [Symbol.toStringTag]: "Temporal.PlainDate";
  readonly year: number;
  readonly month: number;
  readonly day: number;
}

/** A Temporal.PlainDateTime, native or polyfilled; written with its wall-clock time. */
export interface TemporalPlainDateTimeLike {
  readonly [Symbol.toStringTag]: "Temporal.PlainDateTime";
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
  readonly microsecond: number;
  readonly nanosecond: number;
}

/** A Temporal.ZonedDateTime, native or polyfilled; written with its wall-clock time in its zone. */
export interface TemporalZonedDateTimeLike {
  readonly [Symbol.toStringTag]: "Temporal.ZonedDateTime";
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
  readonly millisecond: number;
  readonly microsecond: number;
  readonly nanosecond: number;
}

/** A Temporal.Instant, native or polyfilled; written as UTC. */
export interface TemporalInstantLike {
  readonly [Symbol.toStringTag]: "Temporal.Instant";
  readonly epochMilliseconds: number;
}

export type TemporalLike =
  | TemporalPlainDateLike
  | TemporalPlainDateTimeLike
  | TemporalZonedDateTimeLike
  | TemporalInstantLike;

/**
 * What a cell can hold. `null` and `undefined` write no cell. A `Date` is written as UTC.
 * Temporal values are recognised by their `Symbol.toStringTag`, so no Temporal global is needed.
 */
export type CellValue =
  | string
  | number
  | boolean
  | Date
  | TemporalLike
  | null
  | undefined
  | Registered;

/** A cell written as a plain object, so the value can carry a style. */
export interface CellObject {
  readonly value: CellValue;
  readonly style?: CellStyle;
}

/** One cell of a row: a bare value, or a plain `{ value }` object. */
export type Cell = CellValue | CellObject;

/** One row, left to right from column A. */
export type Row = readonly Cell[];

/** What a column carries for every cell in it. Positions match the row arrays: index 0 is column A. */
export interface Column {
  /** Width in Excel's character units (the unit the column-width dialog shows), from above 0 to 255. */
  readonly width?: number;
  /** Default style for the column's cells; a cell's own style properties win. */
  readonly style?: CellStyle;
}

/** Rows from the top and columns from the left that stay in place while the sheet scrolls. */
export interface Freeze {
  readonly rows?: number;
  readonly columns?: number;
}

/**
 * One sheet of a workbook.
 *
 * @example
 * const sheet: Sheet = {
 *   name: "Orders",
 *   columns: [undefined, { style: { numberFormat: '#,##0.00 "kr"' } }],
 *   rows: [["Order", "Total"], ["A-1001", 1234.5]],
 * };
 */
export interface Sheet {
  readonly name: string;
  /** Any iterable of rows, so a generator or database cursor can feed a sheet without an array. */
  readonly rows: Iterable<Row>;
  /** Per-column settings, by position; leave a slot `undefined` for a column with none. */
  readonly columns?: readonly (Column | undefined)[];
  /** Freeze panes: the header rows and the leading columns that stay visible. */
  readonly freeze?: Freeze;
  /** Filter dropdowns on row 1 across the used range. */
  readonly autoFilter?: boolean;
}

/**
 * Converts a value the writer does not handle natively, or overrides how it handles a `Date` or a
 * Temporal value. Called once per such value with the value's path; the result is written as if the
 * caller had put it in the cell. Never called for strings, numbers, booleans, `null` or `undefined`.
 *
 * @example
 * const coerce: Coerce = (value) => (typeof value === "bigint" ? Number(value) : value);
 */
export type Coerce = (value: Date | TemporalLike | Registered, path: string) => CellValue;
