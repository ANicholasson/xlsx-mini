/**
 * Every failure the reader reports, as a string literal so callers can switch on it.
 *
 * About the file:
 * - `NOT_XLSX`: not a package at all, or a package that is not a workbook: an old `.xls`, an
 *   encrypted workbook, a CSV, another kind of zip.
 * - `INVALID_PACKAGE`: the package is damaged or lies: a directory out of bounds, two parts with
 *   one name, a local header that disagrees with the directory, a part that inflates to another
 *   size than declared, a wrong CRC-32, a compression method other than stored or deflate.
 * - `INVALID_XML`: a part is not well-formed XML in UTF-8, holds a DOCTYPE, nests deeper or holds
 *   a longer token than any workbook does.
 * - `INVALID_WORKBOOK`: the parts do not make a workbook: the sheet list, a sheet's data or the
 *   shared strings cannot be found or are damaged, a relationship leaves the package, two sheets
 *   share a name.
 * - `INVALID_CELL`: a row or a cell is out of order, a reference is malformed, a cell is of an
 *   unknown type or holds what its type does not, a number is not decimal, a shared string does not
 *   exist, a date is stored as text (`t="d"`). `sheet` and `cell` say where.
 * - `LIMIT_EXCEEDED`: the file asks for more than a limit allows. `limit` names the key of
 *   `ReadLimits` to raise.
 * - `MISSING_HEADER`: the header row lacks a declared column. `header` is its text.
 *
 * About the platform or the call:
 * - `NO_INFLATE`: a part is deflated and the platform has no `DecompressionStream`, or one that
 *   reads a deflate stream cut short as whole and reads on past the end of a gzip member.
 * - `UNKNOWN_SHEET`: `rows` was asked for a name or an index the workbook does not have.
 * - `INVALID_LIMIT`: a limit that is not a positive finite number. `limit` names it.
 *
 * @experimental
 */
export type ReadErrorCode =
  | "NOT_XLSX"
  | "INVALID_PACKAGE"
  | "INVALID_XML"
  | "INVALID_WORKBOOK"
  | "INVALID_CELL"
  | "LIMIT_EXCEEDED"
  | "MISSING_HEADER"
  | "NO_INFLATE"
  | "UNKNOWN_SHEET"
  | "INVALID_LIMIT";

/**
 * How much the reader accepts before it refuses a file. Every limit has a default sized for a file
 * a person uploads; a caller sets one per `openXlsx` call, and no value turns one off: each must be
 * a positive finite number.
 *
 * @example
 * const book = await openXlsx(file, { limits: { cellPositions: 30_000_000 } });
 *
 * @experimental
 */
export interface ReadLimits {
  /**
   * Bytes one operation may inflate. Opening the workbook is one operation, together with the
   * parts read once for the whole workbook (shared strings, styles); every `rows` iteration is one.
   * A part's declared size is charged before a byte is inflated. Default 134 217 728 (128 MiB).
   */
  readonly inflatedBytes?: number;
  /**
   * Cell positions one `rows` iteration may yield, empty cells before a value included.
   * Default 5 000 000.
   */
  readonly cellPositions?: number;
  /**
   * Characters of text one row may hold, counted as its values are placed. A row is held whole, so
   * this bounds the text of the values one `rows` iteration holds at a time. Default 4 194 304.
   */
  readonly rowCharacters?: number;
  /** Entries in the shared-string table. Default 1 000 000. */
  readonly sharedStrings?: number;
  /**
   * Bytes the shared-string table may inflate to, on top of counting against `inflatedBytes`. The
   * table is held in memory as long as the workbook. In trials at the default, it held at most 2.9
   * times this on Node and 4.2 times on Bun, for a million entries of 16 characters with one
   * outside Latin-1, which makes an entry two bytes a character. A measurement, not a bound.
   * Default 33 554 432 (32 MiB).
   */
  readonly sharedStringBytes?: number;
  /**
   * Parts in the package, read or not: every entry of its directory counts. A workbook with one
   * picture per row has about one part per row. Default 100 000.
   */
  readonly parts?: number;
}

type ReadLimit = keyof ReadLimits;

/** Where a refusal happened, as far as the code that refuses knows. */
interface ReadErrorDetail {
  readonly sheet?: string | undefined;
  readonly cell?: string | undefined;
  readonly limit?: ReadLimit | undefined;
  readonly header?: string | undefined;
}

/**
 * The one error class the reader throws. `code` says what was refused; `sheet` and `cell` say
 * where, when that is known. The writer's `XlsxMiniError` is never thrown by the reader, and this
 * class never by the writer.
 *
 * @example
 * try {
 *   const book = await openXlsx(file);
 * } catch (error) {
 *   if (error instanceof XlsxReadError && error.code === "NOT_XLSX") console.log("not a workbook");
 * }
 *
 * @experimental
 */
export class XlsxReadError extends Error {
  override readonly name = "XlsxReadError";
  readonly code: ReadErrorCode;
  /**
   * The name of the sheet the refusal is about, when it is about one: a sheet being read, a sheet
   * the opening refuses, and for `UNKNOWN_SHEET` the name asked for. It is whole, unlike `cell`: a
   * name the workbook lists can be nearly as long as a token, 4 194 304 characters.
   */
  readonly sheet: string | undefined;
  /**
   * Where in the sheet an `INVALID_CELL` is, such as `C41`: the cell; for a row, the first cell of
   * the row; for what stands in a row outside a cell, the position after the row's last cell.
   * `undefined` when the place has no reference, as for a row whose number is not a number.
   *
   * A refused reference is given as the file writes it, not decoded: `A&#49;` stays `A&#49;`. One
   * longer than the longest reference, `XFD1048576`, is cut after its first ten UTF-16 code units,
   * or nine when the tenth is the first half of a surrogate pair, and ends in `…`; the cut may fall
   * inside a character reference. Never longer than 16 code units.
   */
  readonly cell: string | undefined;
  /** For `LIMIT_EXCEEDED` and `INVALID_LIMIT`: the key of `ReadLimits` to look at. */
  readonly limit: ReadLimit | undefined;
  /** For `MISSING_HEADER`: the header text that the header row lacks. */
  readonly header: string | undefined;
  /**
   * Starts at the place in the reader that refused. Whether the caller's frames that led there
   * follow is the runtime's choice: Node and Deno give them, Bun 1.4 does not. A refusal that
   * `rows` names the sheet of keeps the stack it was made with.
   */
  declare stack: string;

  constructor(code: ReadErrorCode, message: string, detail: ReadErrorDetail = {}) {
    super(message);
    this.code = code;
    this.sheet = detail.sheet;
    this.cell = detail.cell;
    this.limit = detail.limit;
    this.header = detail.header;
  }
}

const DEFAULTS = {
  inflatedBytes: 134_217_728,
  cellPositions: 5_000_000,
  rowCharacters: 4_194_304,
  sharedStrings: 1_000_000,
  sharedStringBytes: 33_554_432,
  parts: 100_000,
} as const satisfies Required<ReadLimits>;

const STRUCTURE_SHARE = 1 / 8;

/**
 * Brands a meter. A declared class emits no code. It is a class because no object literal can write
 * a class's private member and a spread drops it, so an object literal where a meter is expected
 * does not compile, with a spread or without. The types stop no more than that: a cast makes a
 * meter, and so does a copy made through a function. `test/boundaries.test.ts` searches read-side
 * code other than this file for the forms a reader would reach for: a cast other than `as const`, a
 * type predicate, an object spread or rest, `structuredClone`, `Reflect`, a function of `Object`
 * that does more than read one or freeze it, and `globalThis` other than as `globalThis.name`. The
 * search guards against habit and proves nothing: a helper in a file outside the read side, an
 * overload whose body returns a literal, and a value typed `any` all pass it.
 */
declare class Metered {
  private readonly metered: true;
}

/**
 * The running count of one limit within one operation. `charge` is the only way to learn whether
 * something fits, so no module makes its own comparison or words `LIMIT_EXCEEDED` a second way.
 * Whatever inflates, allocates per position or counts entries needs a meter handed to it, and only
 * `resolveLimits` makes one.
 */
export interface Meter<Limit extends ReadLimit> extends Metered {
  readonly limit: Limit;
  /**
   * Takes `amount` from what is left and from every meter this one is nested in. Throws
   * `LIMIT_EXCEEDED` naming `limit`, and takes nothing, when it does not fit or is not a
   * non-negative number: charge before the work, not after.
   */
  readonly charge: (amount: number) => void;
}

/**
 * What a part's bytes are read for, and the limit that pays for them: the parts that become lists
 * in memory, the shared-string table, a worksheet.
 */
const PART_LIMIT = {
  structure: "inflatedBytes",
  sharedStrings: "sharedStringBytes",
  worksheet: "inflatedBytes",
} as const satisfies Record<string, ReadLimit>;

export type PartUse = keyof typeof PART_LIMIT;

/**
 * A meter of inflated bytes that says which kind of part it pays for. `in out` holds a meter to
 * its one use: it is no meter of "this use or that", so a part whose name leaves its use open
 * between two has no meter to be opened with.
 */
export interface PartMeter<in out Use extends PartUse> extends Meter<(typeof PART_LIMIT)[Use]> {
  readonly use: Use;
}

/** A meter that starts again for every row: what one row may hold, not what a sheet may. */
export interface RowMeter<Limit extends ReadLimit> extends Meter<Limit> {
  /** Starts the next row: what is left returns to the whole allowance. */
  next(): void;
}

/** The meters of the parts read once per workbook. They live as long as the opened workbook. */
interface BookMeters {
  /** One charge per entry of the package's directory, before the entry is kept. */
  readonly parts: Meter<"parts">;
  /**
   * For a part that becomes a list in memory (relationships, the sheet list, styles), for the
   * package's directory and for the part names the relationships resolve to: an eighth of
   * `inflatedBytes`, because a list item can cost twenty times its XML.
   */
  readonly structure: PartMeter<"structure">;
  /** For the shared-string table's part. */
  readonly stringBytes: PartMeter<"sharedStrings">;
  /** One charge per entry of the shared-string table. */
  readonly strings: Meter<"sharedStrings">;
}

/** The meters of one `rows` iteration. Fresh for every iteration, so a sheet can be read again. */
interface RowsMeters {
  /** For the worksheet part. */
  readonly inflate: PartMeter<"worksheet">;
  /** One charge per value placed: the positions it adds to its row, the gap before it included. */
  readonly positions: Meter<"cellPositions">;
  /** One charge per text value placed, of its length; started again for every row. */
  readonly rowText: RowMeter<"rowCharacters">;
}

/**
 * A caller's limits, validated and turned into meters. Holds no readable number: what is charged
 * to what is decided in `resolveLimits` and nowhere else.
 */
interface Limits {
  readonly book: BookMeters;
  /** New meters for one `rows` iteration. */
  rows(): RowsMeters;
}

function allowanceOf(limits: ReadLimits | undefined, limit: ReadLimit): number {
  const value = limits?.[limit];
  if (value === undefined) return DEFAULTS[limit];
  if (Number.isFinite(value) && value > 0) return value;
  throw new XlsxReadError("INVALID_LIMIT", `limits.${limit} must be a positive finite number`, {
    limit,
  });
}

/**
 * `charge` takes from what is left of `allowance` and from `within`, all or nothing. Every meter has
 * `next` at run time; its type decides whether its holder may call it.
 */
function meter<Limit extends ReadLimit>(
  limit: Limit,
  allowance: number,
  within?: Meter<"inflatedBytes">,
): RowMeter<Limit> {
  let left = allowance;
  return {
    limit,
    charge(amount: number): void {
      const fits = amount >= 0 && amount <= left;
      if (!fits) {
        throw new XlsxReadError("LIMIT_EXCEEDED", `more than limits.${limit} allows`, { limit });
      }
      within?.charge(amount);
      left -= amount;
    },
    next(): void {
      left = allowance;
    },
  } as RowMeter<Limit>;
}

function partMeter<Use extends PartUse>(
  use: Use,
  allowance: number,
  within?: Meter<"inflatedBytes">,
): PartMeter<Use> {
  return Object.assign(meter(PART_LIMIT[use], allowance, within), { use });
}

/**
 * Validates the caller's limits over the defaults and turns them into meters. Throws
 * `INVALID_LIMIT` for a value that is not a positive finite number (`Infinity`, `NaN`, `0`, a
 * negative), so no value turns a limit off. Called once per `openXlsx`.
 */
export function resolveLimits(limits: ReadLimits | undefined): Limits {
  const inflatedBytes = allowanceOf(limits, "inflatedBytes");
  const cellPositions = allowanceOf(limits, "cellPositions");
  const rowCharacters = allowanceOf(limits, "rowCharacters");
  const opening = meter("inflatedBytes", inflatedBytes);
  return {
    book: {
      parts: meter("parts", allowanceOf(limits, "parts")),
      structure: partMeter("structure", inflatedBytes * STRUCTURE_SHARE, opening),
      stringBytes: partMeter("sharedStrings", allowanceOf(limits, "sharedStringBytes"), opening),
      strings: meter("sharedStrings", allowanceOf(limits, "sharedStrings")),
    },
    rows: () => ({
      inflate: partMeter("worksheet", inflatedBytes),
      positions: meter("cellPositions", cellPositions),
      rowText: meter("rowCharacters", rowCharacters),
    }),
  };
}
