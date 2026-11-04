import { XlsxMiniError } from "../errors/index.js";
import type { CellStyle } from "../styles/index.js";
import { XML_DECLARATION } from "../xml/index.js";
import type { Sheet } from "./lib/model.js";
import { renderSheetData, type WorksheetContext } from "./lib/render-rows.js";

export type {
  Cell,
  CellObject,
  CellValue,
  Coerce,
  Column,
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

/** Renders `xl/worksheets/sheetN.xml`, interning text and registering styles as it goes. */
export function renderWorksheet(sheet: Sheet, context: WorksheetContext): string {
  const columnStyles = collectColumnStyles(sheet, context);
  const { rows, dimension } = renderSheetData(sheet.rows, columnStyles, context);
  const sheetData = rows.length === 0 ? "<sheetData/>" : `<sheetData>${rows}</sheetData>`;
  return (
    `${XML_DECLARATION}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    dimension +
    sheetData +
    "</worksheet>"
  );
}

/** Column styles by position, validated up front so a bad format is reported at the column. */
function collectColumnStyles(
  sheet: Sheet,
  context: WorksheetContext,
): readonly (CellStyle | undefined)[] {
  const columns = sheet.columns;
  if (columns === undefined) return [];
  const styles: (CellStyle | undefined)[] = [];
  for (let i = 0; i < columns.length; i++) {
    const style = columns[i]?.style;
    if (style !== undefined) {
      try {
        context.styles.validate(style);
      } catch (error) {
        if (error instanceof XlsxMiniError) {
          throw new XlsxMiniError(
            error.code,
            `sheets[${context.sheetIndex}].columns[${i}].style.numberFormat`,
            error.message.replace(/ at $/, ""),
          );
        }
        throw error;
      }
    }
    styles.push(style);
  }
  return styles;
}
