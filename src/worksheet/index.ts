import type { StringTable } from "../shared-strings/index.js";
import { XML_DECLARATION } from "../xml/index.js";
import type { Sheet } from "./lib/model.js";
import { renderSheetData } from "./lib/render-rows.js";

export type { Cell, CellValue, Row, Sheet } from "./lib/model.js";
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
