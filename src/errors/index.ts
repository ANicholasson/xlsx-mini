/** Every failure the writer reports, as a string literal so callers can switch on it. */
export type ErrorCode =
  | "NO_SHEETS"
  | "INVALID_SHEET_NAME"
  | "DUPLICATE_SHEET_NAME"
  | "INVALID_NUMBER"
  | "INVALID_DATE"
  | "DATE_OUT_OF_RANGE"
  | "UNSUPPORTED_VALUE"
  | "INVALID_NUMBER_FORMAT"
  | "INVALID_COLUMN_WIDTH"
  | "INVALID_FREEZE"
  | "CELL_TEXT_TOO_LONG"
  | "TOO_MANY_COLUMNS"
  | "TOO_MANY_ROWS"
  | "PACKAGE_TOO_LARGE"
  | "NO_DOM";

/**
 * The one error class the writer throws. `code` says what went wrong, `path` says where in the
 * input, as the caller wrote it: `sheets[1].rows[40][2]` is the third cell of the 41st row of the
 * second sheet.
 *
 * @example
 * try {
 *   await writeXlsx(workbook);
 * } catch (error) {
 *   if (error instanceof XlsxMiniError && error.code === "INVALID_NUMBER") console.log(error.path);
 * }
 */
export class XlsxMiniError extends Error {
  override readonly name = "XlsxMiniError";
  readonly code: ErrorCode;
  readonly path: string;

  constructor(code: ErrorCode, path: string, message: string) {
    super(`${message} at ${path}`);
    this.code = code;
    this.path = path;
  }
}
