import type { Sheet, StreamedSheet } from "../../worksheet/index.js";

/**
 * Everything `writeXlsx` needs: the sheets, in tab order.
 *
 * @example
 * const workbook: Workbook = { sheets: [{ name: "Sheet1", rows: [["Hello", 42]] }] };
 */
export interface Workbook {
  readonly sheets: readonly Sheet[];
}

/**
 * Everything `streamXlsx` needs: the sheets, in tab order, whose rows may arrive asynchronously.
 * A `Workbook` is a `StreamedWorkbook`.
 *
 * @example
 * const workbook: StreamedWorkbook = { sheets: [{ name: "Orders", rows: cursor.rows() }] };
 */
export interface StreamedWorkbook {
  readonly sheets: readonly StreamedSheet[];
}
