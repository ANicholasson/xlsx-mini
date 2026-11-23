/**
 * `xlsx-mini/read`: the values of an `.xlsx` file, read lazily. The application service of the
 * read side: it opens the package, asks each feature's read side for its part, and owns nothing
 * about any format. Nothing the writer imports may import this file.
 */
import { type ReadLimits, XlsxReadError, resolveLimits } from "./errors/read.js";
import { openPackage, partName } from "./opc/read.js";
import { type SharedStrings, sharedStringsVisitor } from "./shared-strings/read.js";
import { readWorkbook } from "./workbook/read.js";
import { type ReadRow, worksheetVisitor } from "./worksheet/read.js";
import { readXml, scanXml } from "./xml/read.js";

export { type ReadErrorCode, type ReadLimits, XlsxReadError } from "./errors/read.js";
export type { ReadRow, ReadValue } from "./worksheet/read.js";

/**
 * How to read, as opposed to what to read.
 *
 * @example
 * const book = await openXlsx(file, { limits: { parts: 200_000 } });
 *
 * @experimental
 */
export interface ReadOptions {
  /** Limits to use in place of their defaults, for this workbook; see `ReadLimits`. */
  readonly limits?: ReadLimits;
}

/**
 * A workbook as `openXlsx` gives it: the names of its sheets, and each sheet's rows read on
 * demand. It holds the file it was opened from and no other resource, so there is nothing to
 * close.
 *
 * @example
 * const book = await openXlsx(file);
 * console.log(book.sheets); // ["Orders", "Customers"]
 *
 * @experimental
 */
export interface OpenedWorkbook {
  /**
   * The name of every sheet in the order of the workbook, hidden ones included, as the file states
   * it: nothing trimmed or normalised. Never empty, and no two names are the same character for
   * character; names that differ only by case are two sheets.
   */
  readonly sheets: readonly string[];
  /**
   * The rows of one sheet, chosen by its exact name or by its index in `sheets` from 0. Throws
   * `UNKNOWN_SHEET` here, at the call, for a sheet the workbook does not have. Nothing is read
   * until the result is iterated. A chart sheet, a dialog sheet and a macro sheet have no rows.
   *
   * Every iteration reads the sheet from its start, yields each row that holds a value in the
   * order of the sheet, and has limits of its own: two sheets may be read at once, and a sheet
   * again. The shared strings are read once, by the first iteration of a worksheet, charged to the
   * opening's `inflatedBytes`, and held as long as the workbook. Leaving the loop early stops the
   * inflate. A refusal of the sheet names it in `sheet`; a refusal of the shared strings names none,
   * and every later iteration of a worksheet repeats it.
   *
   * One iteration holds one row and the text of the part around it. In trials with the default
   * limits, the most one iteration reached was 90 MiB of heap on Node, for a row of six cells that
   * each hold the longest text of character references, and 309 MiB resident on Bun, for 30 rows
   * that each hold one cell of 4 194 304 carriage returns. These are measurements, not bounds.
   * Iterations at the same time each hold their own, and the file decides how many sheets there
   * are.
   *
   * A sheet is held to its size and its checksum at its last byte, after its rows were yielded:
   * act on what was read once the loop has ended without an error.
   *
   * @example
   * for await (const row of book.rows("Orders")) console.log(row.number, row.cells);
   */
  rows(sheet: string | number): AsyncIterable<ReadRow>;
}

/**
 * Opens an `.xlsx` file for reading its values: reads the package's directory and the sheet list,
 * and nothing else. A `Blob` is read in slices and never whole; bytes are not copied, so they must
 * not change while the workbook is in use. A part the reader has no use for, such as a picture, is
 * never inflated. The search for the end of the file reads up to its last 65 633 bytes, which in a
 * small file hold other parts' bytes too.
 *
 * Built for files from strangers: every refusal is an `XlsxReadError` with a code from
 * `ReadErrorCode`, and a damaged workbook is refused, never repaired and never read as an empty one.
 * A file that is not a workbook package is `NOT_XLSX`. A sheet list that is missing, empty or stated
 * twice is `INVALID_WORKBOOK`, and so are a sheet whose relationship is missing or not to a kind of
 * sheet, a target outside the package, and two sheets with one name.
 *
 * @example
 * const book = await openXlsx(file);
 * console.log(book.sheets);
 *
 * @experimental
 */
export async function openXlsx(
  file: Uint8Array | ArrayBuffer | Blob,
  options?: ReadOptions,
): Promise<OpenedWorkbook> {
  const limits = resolveLimits(options?.limits);
  const { book } = limits;
  const pkg = await openPackage(file, book.parts, book.structure);
  const { sheets: listed, sharedStrings } = await readWorkbook(
    (name) => pkg.open(partName(name, "structure"), book.structure),
    book.structure,
  );
  const named = new Map(listed.map((listing) => [listing.name, listing] as const));

  // A rejected promise that nothing awaits ends the process as an unhandled rejection.
  let strings: Promise<SharedStrings> | undefined;
  async function readStrings(): Promise<SharedStrings> {
    if (sharedStrings === undefined) return [];
    const bytes =
      pkg.open(sharedStrings, book.stringBytes) ?? damaged("the shared-string table is missing");
    return readXml(bytes, sharedStringsVisitor(book.strings));
  }

  return {
    sheets: Object.freeze(listed.map(({ name }) => name)),
    rows(sheet) {
      const found = typeof sheet === "number" ? listed[sheet] : named.get(sheet);
      if (found === undefined) {
        throw new XlsxReadError(
          "UNKNOWN_SHEET",
          "the workbook has no such sheet",
          typeof sheet === "string" ? { sheet } : {},
        );
      }
      const { name, worksheet } = found;
      return {
        async *[Symbol.asyncIterator]() {
          if (worksheet === undefined) return;
          const table = await (strings ??= readStrings());
          const meters = limits.rows();
          try {
            const bytes =
              pkg.open(worksheet, meters.inflate) ?? damaged("a sheet's part is missing");
            yield* scanXml(
              bytes,
              worksheetVisitor({
                strings: table,
                positions: meters.positions,
                rowText: meters.rowText,
              }),
            );
          } catch (error) {
            if (!(error instanceof XlsxReadError)) throw error;
            const named = new XlsxReadError(error.code, error.message, {
              sheet: name,
              cell: error.cell,
              limit: error.limit,
              header: error.header,
            });
            named.stack = error.stack;
            throw named;
          }
        },
      };
    },
  };
}

function damaged(message: string): never {
  throw new XlsxReadError("INVALID_WORKBOOK", message);
}
