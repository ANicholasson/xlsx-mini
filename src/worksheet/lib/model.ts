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

/** A cell written as a plain object, for when the value needs company (a style, later). */
export interface CellObject {
  readonly value: CellValue;
}

/** One cell of a row: a bare value, or a plain `{ value }` object. */
export type Cell = CellValue | CellObject;

/** One row, left to right from column A. */
export type Row = readonly Cell[];

/**
 * One sheet of a workbook.
 *
 * @example
 * const sheet: Sheet = { name: "Orders", rows: [["Order", "Total"], ["A-1001", 1234.5]] };
 */
export interface Sheet {
  readonly name: string;
  /** Any iterable of rows, so a generator or database cursor can feed a sheet without an array. */
  readonly rows: Iterable<Row>;
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
