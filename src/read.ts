/**
 * `xlsx-mini/read`: the values of an `.xlsx` file, read lazily. The application service of the
 * read side: it opens the package, asks each feature's read side for its part, and owns nothing
 * about any format. Nothing the writer imports may import this file.
 */
import { type ReadLimits, XlsxReadError, resolveLimits } from "./errors/read.js";
import { openPackage, partName } from "./opc/read.js";
import { readWorkbook } from "./workbook/read.js";

export { type ReadErrorCode, type ReadLimits, XlsxReadError } from "./errors/read.js";

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
   * Reading the rows of a worksheet is not built yet: iterating them rejects with an `Error` that
   * says so, and never ends as if the sheet were empty.
   *
   * @example
   * for await (const row of book.rows("Orders")) console.log(row);
   */
  rows(sheet: string | number): AsyncIterable<never>;
}

/**
 * Opens an `.xlsx` file for reading its values: reads the package's directory and the sheet list,
 * and nothing else. A `Blob` is read in slices and never whole; bytes are not copied, so they must
 * not change while the workbook is in use. Parts the reader has no use for, such as pictures, are
 * never read.
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
  options: ReadOptions = {},
): Promise<OpenedWorkbook> {
  const { book } = resolveLimits(options.limits);
  const pkg = await openPackage(file, book.parts, book.structure);
  const listed = await readWorkbook((name) =>
    pkg.open(partName(name, "structure"), book.structure),
  );
  return {
    sheets: listed.map(({ name }) => name),
    rows(sheet) {
      const found =
        typeof sheet === "number" ? listed[sheet] : listed.find(({ name }) => name === sheet);
      if (found === undefined) {
        throw new XlsxReadError("UNKNOWN_SHEET", "the workbook has no such sheet");
      }
      return {
        async *[Symbol.asyncIterator]() {
          if (found.worksheet) throw new Error("reading the rows of a worksheet is not built yet");
          yield* [];
        },
      };
    },
  };
}
