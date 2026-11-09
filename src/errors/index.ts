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
  | "INVALID_STYLE"
  | "INVALID_COLUMN_WIDTH"
  | "INVALID_FREEZE"
  | "INVALID_MERGE"
  | "INVALID_ROW_HEIGHT"
  | "CELL_TEXT_TOO_LONG"
  | "TOO_MANY_COLUMNS"
  | "TOO_MANY_ROWS"
  | "PACKAGE_TOO_LARGE"
  | "ASYNC_ROWS"
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
    super(path.length === 0 ? message : `${message} at ${path}`);
    this.code = code;
    this.path = path;
  }
}

/**
 * The same error at a longer path: a feature that validates without knowing where its input sits
 * throws with the local path, and the caller prefixes its own. `prefix` ends up as `prefix.local`,
 * or as `prefix` alone when the local path is empty.
 */
export function atPath(error: XlsxMiniError, prefix: string): XlsxMiniError {
  const suffix = ` at ${error.path}`;
  const message = error.message.endsWith(suffix)
    ? error.message.slice(0, -suffix.length)
    : error.message;
  return new XlsxMiniError(
    error.code,
    error.path.length === 0 ? prefix : `${prefix}.${error.path}`,
    message,
  );
}

/**
 * A value as an error message names it: strings quoted, other primitives as they print, objects
 * by their tag or constructor. Never throws, unlike `JSON.stringify` on a bigint.
 */
export function describeValue(value: unknown): string {
  if (typeof value === "string") return JSON.stringify(value);
  if (typeof value === "bigint") return `${value}n`;
  if (typeof value !== "object" || value === null) return String(value);
  const tag = (value as { readonly [Symbol.toStringTag]?: unknown })[Symbol.toStringTag];
  if (typeof tag === "string") return tag;
  const constructor = (value as { constructor?: { name?: string } }).constructor;
  return constructor?.name ?? "object";
}
