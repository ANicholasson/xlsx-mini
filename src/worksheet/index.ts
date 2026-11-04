import { XML_DECLARATION } from "../xml/index.js";
import type { Sheet } from "./lib/model.js";
import { renderSheetData, type WorksheetContext } from "./lib/render-rows.js";

export type {
  Cell,
  CellObject,
  CellValue,
  Coerce,
  Register,
  Registered,
  Row,
  Sheet,
  TemporalInstantLike,
  TemporalLike,
  TemporalPlainDateLike,
  TemporalPlainDateTimeLike,
  TemporalZonedDateTimeLike,
} from "./lib/model.js";
export type { WorksheetContext } from "./lib/render-rows.js";
export { columnLetter } from "./lib/cell-ref.js";
export { serialFromCivil, serialFromEpochMillis } from "./lib/serial.js";

/** Renders `xl/worksheets/sheetN.xml`, interning text and allocating date styles as it goes. */
export function renderWorksheet(sheet: Sheet, context: WorksheetContext): string {
  const { rows, dimension } = renderSheetData(sheet.rows, context);
  const sheetData = rows.length === 0 ? "<sheetData/>" : `<sheetData>${rows}</sheetData>`;
  return (
    `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    dimension +
    sheetData +
    "</worksheet>"
  );
}
