import { XML_DECLARATION } from "../xml/index.js";
import { collectColumns, renderFreeze } from "./lib/layout.js";
import type { Sheet } from "./lib/model.js";
import { renderSheetData, type WorksheetContext } from "./lib/render-rows.js";

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
} from "./lib/model.js";
export type { WorksheetContext } from "./lib/render-rows.js";
export { columnLetter } from "./lib/cell-ref.js";
export {
  type ObjectColumn,
  type SheetFromObjectsOptions,
  sheetFromObjects,
} from "./lib/from-objects.js";
export { serialFromCivil, serialFromEpochMillis } from "./lib/serial.js";

/** A rendered worksheet part and what the workbook part needs to know about it. */
export interface RenderedWorksheet {
  readonly xml: string;
  /** The autofilter range such as `A1:C4`, when the sheet filters a non-empty range. */
  readonly filterRange: string | undefined;
}

/** Renders `xl/worksheets/sheetN.xml`, interning text and registering styles as it goes. */
export function renderWorksheet(sheet: Sheet, context: WorksheetContext): RenderedWorksheet {
  const columns = collectColumns(sheet, context.sheetIndex, (style) =>
    context.styles.validate(style),
  );
  const views = renderFreeze(sheet, context.sheetIndex);
  const { rows, range } = renderSheetData(sheet.rows, columns.styles, context);
  const dimension = range.length === 0 ? "" : `<dimension ref="${range}"/>`;
  const sheetData = rows.length === 0 ? "<sheetData/>" : `<sheetData>${rows}</sheetData>`;
  const filterRange = sheet.autoFilter === true && range.length > 0 ? range : undefined;
  const autoFilter = filterRange === undefined ? "" : `<autoFilter ref="${filterRange}"/>`;
  const xml =
    `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    dimension +
    views +
    columns.cols +
    sheetData +
    autoFilter +
    "</worksheet>";
  return { xml, filterRange };
}
