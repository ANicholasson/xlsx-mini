export { type ErrorCode, XlsxMiniError } from "./errors/index.js";
export type { CompressionMode } from "./opc/index.js";
export {
  type ObjectColumn,
  type SheetFromObjectsOptions,
  sheetFromObjects,
} from "./worksheet/index.js";
export { type WriteOptions, writeXlsx } from "./write-xlsx.js";
export type { CellStyle } from "./styles/index.js";
export type { Workbook } from "./workbook/index.js";
export type {
  Cell,
  CellObject,
  CellValue,
  Coerce,
  Column,
  Freeze,
  Register,
  Registered,
  Row,
  Sheet,
  TemporalInstantLike,
  TemporalLike,
  TemporalPlainDateLike,
  TemporalPlainDateTimeLike,
  TemporalZonedDateTimeLike,
} from "./worksheet/index.js";

/**
 * The media type of an .xlsx file, for `Content-Type` headers and `Blob` construction.
 *
 * @example
 * new Response(bytes, { headers: { "Content-Type": XLSX_MIME_TYPE } })
 */
export const XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
