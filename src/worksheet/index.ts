import type { StringTable } from "../shared-strings/index.js";
import { XML_DECLARATION } from "../xml/index.js";
import { renderSheetData } from "./lib/render-rows.js";

/** What a cell can hold. Grows with the cell-model work (booleans, dates, Temporal). */
export type CellValue = string | number | null | undefined;

/** One cell of a row. `null` and `undefined` write no cell but keep the column position. */
export type Cell = CellValue;

/** One row, left to right from column A. */
export type Row = readonly Cell[];

/**
 * One worksheet.
 *
 * @example
 * const sheet: Sheet = { name: "Orders", rows: [["Order", "Total"], ["A-1001", 1234.5]] };
 */
export interface Sheet {
  readonly name: string;
  /** Any iterable of rows, so a generator or database cursor can feed a sheet without an array. */
  readonly rows: Iterable<Row>;
}

export { columnLetter } from "./lib/cell-ref.js";

/** Renders `xl/worksheets/sheetN.xml`, interning text into the shared-strings table as it goes. */
export function renderWorksheet(sheet: Sheet, strings: StringTable): string {
  const { rows, dimension } = renderSheetData(sheet.rows, strings);
  const sheetData = rows.length === 0 ? "<sheetData/>" : `<sheetData>${rows}</sheetData>`;
  return (
    `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    dimension +
    sheetData +
    "</worksheet>"
  );
}
