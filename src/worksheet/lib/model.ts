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

/**
 * A cell holding a formula, written without the leading `=`; the applications compute it on open.
 * `value` is an optional cached result, shown until then. Formulas are escaped, never evaluated.
 *
 * @example
 * const total: FormulaCell = { formula: "SUM(B2:B4)", style: { bold: true } };
 */
export interface FormulaCell {
  readonly formula: string;
  readonly value?: CellValue;
  readonly style?: CellStyle;
}

/**
 * A cell whose text opens an external target: an absolute URL such as `https://`, `mailto:` or
 * `file:`. The target goes into the sheet's relationships part.
 *
 * @example
 * const docs: LinkCell = { value: "Documentation", href: "https://example.com/docs" };
 */
export interface LinkCell {
  readonly value: CellValue;
  readonly href: string;
  readonly style?: CellStyle;
}

/**
 * A cell whose text jumps to a cell or range in this workbook: in the same sheet when `sheet` is
 * left out, otherwise in the named sheet.
 *
 * @example
 * const back: SheetLinkCell = { value: "Back to summary", sheet: "Summary", cell: "A1" };
 */
export interface SheetLinkCell {
  readonly value: CellValue;
  /** The sheet to jump to, by name; the current sheet when left out. */
  readonly sheet?: string;
  /** A cell such as `B2` or a range such as `A1:C4`. */
  readonly cell: string;
  readonly style?: CellStyle;
}

/**
 * A cell showing a PNG, anchored to the cell at its pixel size or the size given; one of `width`
 * and `height` scales the other. The row grows to the picture, and so does a column without a
 * numeric width. The cell beneath stays empty; `style` reaches it for a fill or a border.
 *
 * @example
 * const logo: ImageCell = { image: pngBytes, height: 40 };
 */
export interface ImageCell {
  readonly image: Uint8Array;
  readonly width?: number;
  readonly height?: number;
  readonly style?: CellStyle;
}

/** One cell of a row: a bare value, or a plain object carrying a value, a formula, a link or an image. */
export type Cell = CellValue | CellObject | FormulaCell | LinkCell | SheetLinkCell | ImageCell;

/** One row, left to right from column A. */
export type Row = readonly Cell[];

/** What a column carries for every cell in it. Positions match the row arrays: index 0 is column A. */
export interface Column {
  /**
   * Width in Excel's character units (the unit the column-width dialog shows), from above 0 to
   * 255, or `"auto"` to fit the column to its widest cell, from 3 to 60. When streaming, the first
   * 1 000 rows set an auto width.
   */
  readonly width?: number | "auto";
  /** Default style for the column's cells; a cell's own style properties win. */
  readonly style?: CellStyle;
}

/** Rows from the top and columns from the left that stay in place while the sheet scrolls. */
export interface Freeze {
  readonly rows?: number;
  readonly columns?: number;
}

/** A message shown while a validated cell is selected. */
export interface ValidationMessage {
  /** At most 32 characters. */
  readonly title?: string;
  /** At most 255 characters. */
  readonly message: string;
}

/** The message shown when an entry fails validation, and whether Excel refuses the entry. */
export interface ValidationError extends ValidationMessage {
  /** `"stop"` (the default) refuses the entry, `"warning"` asks, `"information"` only tells. */
  readonly style?: "stop" | "warning" | "information";
}

/** Bounds on a number: with both ends a range, with one end a minimum or a maximum. */
export interface NumberBounds {
  readonly min?: number;
  readonly max?: number;
  /** Accept whole numbers only. */
  readonly integer?: boolean;
}

/** Bounds on a date, each written as UTC like a `Date` in a cell. */
export interface DateBounds {
  readonly min?: Date;
  readonly max?: Date;
}

/** What every validation rule shares: where it applies and what it says. */
export interface ValidationBase {
  /** A cell such as `B2`, a range such as `B2:B100`, or whole columns such as `B:B`. */
  readonly range: string;
  /** Accept an empty cell; true unless set to false, as in Excel's dialog. */
  readonly allowBlank?: boolean;
  readonly input?: ValidationMessage;
  readonly error?: ValidationError;
}

/**
 * A rule on the cells of a range: a dropdown from a list of values or from a range of cells, a
 * number or a date within bounds. Exactly one of `list`, `number` and `date` is set.
 *
 * @example
 * const status: DataValidation = { range: "C2:C1000", list: ["Open", "Closed"] };
 * const quantity: DataValidation = {
 *   range: "D2:D1000",
 *   number: { min: 1, integer: true },
 *   error: { message: "Enter a whole number of at least 1" },
 * };
 */
export type DataValidation = ValidationBase &
  (
    | {
        /** The values offered, or a range they are read from such as `Lists!A2:A9`. */
        readonly list: readonly string[] | string;
      }
    | { readonly number: NumberBounds }
    | { readonly date: DateBounds }
  );

/** Paper sizes by name, written as Excel's paper size codes. */
export type PaperSize = "letter" | "legal" | "tabloid" | "A3" | "A4" | "A5";

/** Page margins in inches, as Excel stores them; a side left out keeps Excel's default. */
export interface PrintMargins {
  readonly top?: number;
  readonly bottom?: number;
  readonly left?: number;
  readonly right?: number;
  readonly header?: number;
  readonly footer?: number;
}

/**
 * How the sheet prints. Header and footer text uses Excel's codes: `&L`, `&C` and `&R` start the
 * left, centre and right sections, `&P` is the page number, `&N` the page count, `&D` the date,
 * `&T` the time, `&F` the file name, `&A` the sheet name, and `&&` a literal ampersand.
 *
 * @example
 * const print: PrintSetup = {
 *   orientation: "landscape",
 *   paperSize: "A4",
 *   fitTo: { width: 1 },
 *   repeatRows: 1,
 *   footer: "&L&A&RPage &P of &N",
 * };
 */
export interface PrintSetup {
  readonly orientation?: "portrait" | "landscape";
  readonly paperSize?: PaperSize;
  /** Shrink to fit this many pages wide or tall; a dimension left out takes as many pages as needed. */
  readonly fitTo?: { readonly width?: number; readonly height?: number };
  readonly margins?: PrintMargins;
  /** At most 255 characters each. */
  readonly header?: string;
  readonly footer?: string;
  /** The number of rows from the top repeated at the top of every printed page. */
  readonly repeatRows?: number;
}

/**
 * One sheet of a workbook as `streamXlsx` takes it: rows may arrive asynchronously, so a database
 * cursor can feed the sheet page by page while the package streams out.
 *
 * @example
 * const sheet: StreamedSheet = { name: "Orders", rows: cursor.rows() };
 */
export interface StreamedSheet {
  readonly name: string;
  /** Any iterable or async iterable of rows. */
  readonly rows: Iterable<Row> | AsyncIterable<Row>;
  /** Per-column settings, by position; leave a slot `undefined` for a column with none. */
  readonly columns?: readonly (Column | undefined)[];
  /** Freeze panes: the header rows and the leading columns that stay visible. */
  readonly freeze?: Freeze;
  /** Filter dropdowns on row 1 across the used range. */
  readonly autoFilter?: boolean;
  /** Fit every column without a numeric width to its widest cell, as `width: "auto"` does for one. */
  readonly autoWidth?: boolean;
  /** Ranges to merge into one cell each, in A1 form such as `"A1:D1"`; the top-left cell keeps its value. */
  readonly merges?: readonly string[];
  /** Row heights in points by row number, counted from 1: `{ 1: 28 }` makes the first row 28 points. */
  readonly rowHeights?: Readonly<Record<number, number>>;
  /** Rules on what the cells of a range accept, with their dropdowns and messages. */
  readonly validations?: readonly DataValidation[];
  /** How the sheet prints: orientation, paper, fit, margins, header and footer, repeated rows. */
  readonly print?: PrintSetup;
}

/**
 * One sheet of a workbook.
 *
 * @example
 * const sheet: Sheet = {
 *   name: "Orders",
 *   columns: [undefined, { style: { numberFormat: '#,##0.00 "kr"' } }],
 *   rows: [["Order", "Total"], ["A-1001", 1234.5]],
 *   merges: ["A1:B1"],
 * };
 */
export interface Sheet extends StreamedSheet {
  /** Any iterable of rows, so a generator can feed a sheet without an array. */
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
