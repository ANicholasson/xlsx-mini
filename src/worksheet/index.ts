import { XML_DECLARATION, createXmlSink } from "../xml/index.js";
import { collectColumns, collectRowHeights, renderFreeze, renderMerges } from "./lib/layout.js";
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
  readonly data: Uint8Array;
  /** The autofilter range such as `A1:C4`, when the sheet filters a non-empty range. */
  readonly filterRange: string | undefined;
}

/** Renders `xl/worksheets/sheetN.xml`, interning text and registering styles as it goes. */
export function renderWorksheet(sheet: Sheet, context: WorksheetContext): RenderedWorksheet {
  const columns = collectColumns(sheet, context.sheetIndex, (style) =>
    context.styles.validate(style),
  );
  const views = renderFreeze(sheet, context.sheetIndex);
  const merges = renderMerges(sheet, context.sheetIndex);
  const rowHeights = collectRowHeights(sheet, context.sheetIndex);
  const sink = createXmlSink();
  const { range, hasRows } = renderSheetData(sheet.rows, columns.styles, context, sink, rowHeights);
  const noCells = range.length === 0;
  const dimension = noCells ? "" : `<dimension ref="${range}"/>`;
  const filterRange = sheet.autoFilter === true && !noCells ? range : undefined;
  const autoFilter = filterRange === undefined ? "" : `<autoFilter ref="${filterRange}"/>`;
  const prefix =
    `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    dimension +
    views +
    columns.cols +
    (hasRows ? "<sheetData>" : "<sheetData/>");
  const suffix = (hasRows ? "</sheetData>" : "") + autoFilter + merges + "</worksheet>";
  return { data: sink.finish(prefix, suffix), filterRange };
}
