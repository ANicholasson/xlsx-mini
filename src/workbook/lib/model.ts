import type { Sheet } from "../../worksheet/index.js";

/**
 * Everything `writeXlsx` needs: the sheets, in tab order.
 *
 * @example
 * const workbook: Workbook = { sheets: [{ name: "Sheet1", rows: [["Hello", 42]] }] };
 */
export interface Workbook {
  readonly sheets: readonly Sheet[];
}
