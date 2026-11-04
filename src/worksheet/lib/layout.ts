import { XlsxMiniError } from "../../errors/index.js";
import type { CellStyle } from "../../styles/index.js";
import { columnLetter } from "./cell-ref.js";
import type { Sheet } from "./model.js";

const MAX_COLUMN_WIDTH = 255;
const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;

/** What the sheet's columns contribute: a style per position and the `<cols>` element. */
export interface CollectedColumns {
  readonly styles: readonly (CellStyle | undefined)[];
  readonly cols: string;
}

/** Validates every column up front, so a bad width or format is reported at the column. */
export function collectColumns(
  sheet: Sheet,
  sheetIndex: number,
  validateStyle: (style: CellStyle) => void,
): CollectedColumns {
  const columns = sheet.columns;
  if (columns === undefined) return { styles: [], cols: "" };
  const styles: (CellStyle | undefined)[] = [];
  let cols = "";
  for (let i = 0; i < columns.length; i++) {
    const column = columns[i];
    const style = column?.style;
    if (style !== undefined) {
      try {
        validateStyle(style);
      } catch (error) {
        if (error instanceof XlsxMiniError) {
          throw new XlsxMiniError(
            error.code,
            `sheets[${sheetIndex}].columns[${i}].style.numberFormat`,
            error.message.replace(/ at $/, ""),
          );
        }
        throw error;
      }
    }
    styles.push(style);
    const width = column?.width;
    if (width !== undefined) {
      if (!Number.isFinite(width) || width <= 0 || width > MAX_COLUMN_WIDTH) {
        throw new XlsxMiniError(
          "INVALID_COLUMN_WIDTH",
          `sheets[${sheetIndex}].columns[${i}].width`,
          `Column widths run from above 0 to ${MAX_COLUMN_WIDTH} characters, not ${width}`,
        );
      }
      cols += `<col min="${i + 1}" max="${i + 1}" width="${width}" customWidth="1"/>`;
    }
  }
  return { styles, cols: cols.length === 0 ? "" : `<cols>${cols}</cols>` };
}

/** The `<sheetViews>` element with a frozen pane, or an empty string when nothing is frozen. */
export function renderFreeze(sheet: Sheet, sheetIndex: number): string {
  const freeze = sheet.freeze;
  if (freeze === undefined) return "";
  const rows = freeze.rows ?? 0;
  const columns = freeze.columns ?? 0;
  if (!Number.isInteger(rows) || rows < 0 || rows >= MAX_ROWS) {
    throw new XlsxMiniError(
      "INVALID_FREEZE",
      `sheets[${sheetIndex}].freeze.rows`,
      `Frozen rows must be a whole number from 0 to ${MAX_ROWS - 1}, not ${rows}`,
    );
  }
  if (!Number.isInteger(columns) || columns < 0 || columns >= MAX_COLUMNS) {
    throw new XlsxMiniError(
      "INVALID_FREEZE",
      `sheets[${sheetIndex}].freeze.columns`,
      `Frozen columns must be a whole number from 0 to ${MAX_COLUMNS - 1}, not ${columns}`,
    );
  }
  if (rows === 0 && columns === 0) return "";
  const topLeft = columnLetter(columns) + (rows + 1);
  const activePane = rows === 0 ? "topRight" : columns === 0 ? "bottomLeft" : "bottomRight";
  const splits =
    (columns === 0 ? "" : ` xSplit="${columns}"`) + (rows === 0 ? "" : ` ySplit="${rows}"`);
  return `<sheetViews><sheetView workbookViewId="0"><pane${splits} topLeftCell="${topLeft}" activePane="${activePane}" state="frozen"/></sheetView></sheetViews>`;
}
