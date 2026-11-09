import { XlsxMiniError, atPath, describeValue } from "../../errors/index.js";
import type { CellStyle } from "../../styles/index.js";
import { columnLetter } from "./cell-ref.js";
import type { StreamedSheet } from "./model.js";

const MAX_COLUMN_WIDTH = 255;
const MAX_ROW_HEIGHT = 409;
const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;
const RANGE = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/;

/** What the sheet's columns contribute: a style per position and the `<cols>` element. */
export interface CollectedColumns {
  readonly styles: readonly (CellStyle | undefined)[];
  readonly cols: string;
}

/** Validates every column up front, so a bad width or style is reported at the column. */
export function collectColumns(
  sheet: StreamedSheet,
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
          throw atPath(error, `sheets[${sheetIndex}].columns[${i}].style`);
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
export function renderFreeze(sheet: StreamedSheet, sheetIndex: number): string {
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

/** Zero-based column index of column letters such as `A`, `Z` or `AA`. */
function columnIndex(letters: string): number {
  let index = 0;
  for (let i = 0; i < letters.length; i++) index = index * 26 + (letters.charCodeAt(i) - 64);
  return index - 1;
}

interface Box {
  readonly ref: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

/**
 * The `<mergeCells>` element, or an empty string. Every range is validated as A1 form inside the
 * sheet, larger than one cell, and overlapping no other: Excel repairs a file whose merges overlap.
 */
export function renderMerges(sheet: StreamedSheet, sheetIndex: number): string {
  const merges = sheet.merges;
  if (merges === undefined || merges.length === 0) return "";
  const boxes: Box[] = [];
  for (let i = 0; i < merges.length; i++) {
    const ref = merges[i]!;
    const path = `sheets[${sheetIndex}].merges[${i}]`;
    const match = typeof ref === "string" ? RANGE.exec(ref) : null;
    if (match === null) {
      throw new XlsxMiniError(
        "INVALID_MERGE",
        path,
        `A merge is a range in A1 form such as "A1:D1", not ${describeValue(ref)}`,
      );
    }
    const left = columnIndex(match[1]!);
    const top = Number(match[2]);
    const right = columnIndex(match[3]!);
    const bottom = Number(match[4]);
    if (right >= MAX_COLUMNS || bottom > MAX_ROWS) {
      throw new XlsxMiniError("INVALID_MERGE", path, `Range ${ref} lies outside the sheet`);
    }
    if (right < left || bottom < top) {
      throw new XlsxMiniError(
        "INVALID_MERGE",
        path,
        `Range ${ref} must run from its top-left cell to its bottom-right cell`,
      );
    }
    if (right === left && bottom === top) {
      throw new XlsxMiniError("INVALID_MERGE", path, `Range ${ref} is a single cell`);
    }
    const box: Box = { ref, top, left, bottom, right };
    for (const other of boxes) {
      if (
        box.left <= other.right &&
        other.left <= box.right &&
        box.top <= other.bottom &&
        other.top <= box.bottom
      ) {
        throw new XlsxMiniError(
          "INVALID_MERGE",
          path,
          `Range ${ref} overlaps ${other.ref}; Excel refuses overlapping merges`,
        );
      }
    }
    boxes.push(box);
  }
  let xml = `<mergeCells count="${boxes.length}">`;
  for (const box of boxes) xml += `<mergeCell ref="${box.ref}"/>`;
  return `${xml}</mergeCells>`;
}

/** Row heights validated up front: whole row numbers from 1, heights from above 0 to 409 points. */
export function collectRowHeights(
  sheet: StreamedSheet,
  sheetIndex: number,
): Readonly<Record<number, number>> | undefined {
  const heights = sheet.rowHeights;
  if (heights === undefined) return undefined;
  for (const key of Object.keys(heights)) {
    const row = Number(key);
    const path = `sheets[${sheetIndex}].rowHeights[${key}]`;
    if (!Number.isInteger(row) || row < 1 || row > MAX_ROWS) {
      throw new XlsxMiniError(
        "INVALID_ROW_HEIGHT",
        path,
        `Row numbers run from 1 to ${MAX_ROWS}, not ${key}`,
      );
    }
    const height = heights[row];
    if (typeof height !== "number" || !(height > 0) || height > MAX_ROW_HEIGHT) {
      throw new XlsxMiniError(
        "INVALID_ROW_HEIGHT",
        path,
        `Row heights run from above 0 to ${MAX_ROW_HEIGHT} points, not ${height}`,
      );
    }
  }
  return heights;
}
