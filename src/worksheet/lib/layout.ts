import { XlsxMiniError, atPath, describeValue } from "../../errors/index.js";
import type { CellStyle } from "../../styles/index.js";
import {
  type Box,
  MAX_COLUMNS,
  MAX_ROWS,
  boxProblem,
  columnLetter,
  parseReference,
} from "./cell-ref.js";
import type { StreamedSheet } from "./model.js";

const MAX_COLUMN_WIDTH = 255;
const MAX_ROW_HEIGHT = 409;

/** What the sheet's columns contribute: a style per position, which columns size themselves, and the `<cols>` element. */
export interface CollectedColumns {
  readonly styles: readonly (CellStyle | undefined)[];
  /** Whether each listed column fits itself to its cells, or undefined when no column does. */
  readonly measured: readonly boolean[] | undefined;
  /** Whether columns past the list fit themselves: the sheet's `autoWidth`. */
  readonly beyond: boolean;
  /**
   * The `<cols>` element from the numeric widths, the measured ones and the widest picture per
   * column in pixels, or an empty string. A numeric width wins; a picture widens any other column.
   */
  readonly cols: (
    measuredWidths: readonly (number | undefined)[] | undefined,
    imagePixels: readonly number[],
  ) => string;
}

/** Validates every column up front, so a bad width or style is reported at the column. */
export function collectColumns(
  sheet: StreamedSheet,
  sheetIndex: number,
  validateStyle: (style: CellStyle) => void,
): CollectedColumns {
  const columns = sheet.columns ?? [];
  const beyond = sheet.autoWidth === true;
  const styles: (CellStyle | undefined)[] = [];
  const fixed: (number | undefined)[] = [];
  const measured: boolean[] = [];
  let anyMeasured = beyond;
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
    const width: unknown = column?.width;
    if (width === "auto") {
      measured.push(true);
      fixed.push(undefined);
      anyMeasured = true;
      continue;
    }
    if (width !== undefined) {
      if (
        typeof width !== "number" ||
        !Number.isFinite(width) ||
        width <= 0 ||
        width > MAX_COLUMN_WIDTH
      ) {
        throw new XlsxMiniError(
          "INVALID_COLUMN_WIDTH",
          `sheets[${sheetIndex}].columns[${i}].width`,
          `Column widths run from above 0 to ${MAX_COLUMN_WIDTH} characters, or "auto", not ${describeValue(width)}`,
        );
      }
    }
    fixed.push(width);
    measured.push(width === undefined && beyond);
  }
  return {
    styles,
    measured: anyMeasured ? measured : undefined,
    beyond,
    cols: (measuredWidths, imagePixels) => {
      let cols = "";
      const count = Math.max(fixed.length, measuredWidths?.length ?? 0, imagePixels.length);
      for (let i = 0; i < count; i++) {
        let width = fixed[i] ?? measuredWidths?.[i];
        const pixels = imagePixels[i];
        if (fixed[i] === undefined && pixels !== undefined && pixels > 0) {
          const forImage = Math.min(MAX_COLUMN_WIDTH, Math.round(((pixels + 12) / 7) * 100) / 100);
          if (width === undefined || forImage > width) width = forImage;
        }
        if (width === undefined) continue;
        cols += `<col min="${i + 1}" max="${i + 1}" width="${width}" customWidth="1"/>`;
      }
      return cols.length === 0 ? "" : `<cols>${cols}</cols>`;
    },
  };
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

/**
 * The sheet's merges as boxes. Every range is validated as A1 form inside the sheet, larger than
 * one cell, and overlapping no other: Excel repairs a file whose merges overlap.
 */
export function collectMerges(sheet: StreamedSheet, sheetIndex: number): readonly Box[] {
  const merges = sheet.merges;
  const boxes: Box[] = [];
  if (merges === undefined) return boxes;
  for (let i = 0; i < merges.length; i++) {
    const ref = merges[i]!;
    const path = `sheets[${sheetIndex}].merges[${i}]`;
    // A merge names its cells: whole columns, which the parser expands, are not accepted.
    const box = parseReference(ref);
    if (box === undefined || box.ref !== ref) {
      throw new XlsxMiniError(
        "INVALID_MERGE",
        path,
        `A merge is a range in A1 form such as "A1:D1", not ${describeValue(ref)}`,
      );
    }
    const problem = boxProblem(box);
    if (problem !== undefined) throw new XlsxMiniError("INVALID_MERGE", path, problem);
    if (box.right === box.left && box.bottom === box.top) {
      throw new XlsxMiniError("INVALID_MERGE", path, `Range ${ref} is a single cell`);
    }
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
  return boxes;
}

/** The `<mergeCells>` element, or an empty string when there are none. */
export function renderMerges(boxes: readonly Box[]): string {
  if (boxes.length === 0) return "";
  let xml = `<mergeCells count="${boxes.length}">`;
  for (const box of boxes) xml += `<mergeCell ref="${box.ref}"/>`;
  return `${xml}</mergeCells>`;
}

/** The cells that start a merge across more than one column, keyed as the measurer expects. */
export function mergeStarts(boxes: readonly Box[]): ReadonlySet<number> | undefined {
  let starts: Set<number> | undefined;
  for (const box of boxes) {
    if (box.right === box.left) continue;
    starts ??= new Set();
    starts.add((box.top - 1) * MAX_COLUMNS + box.left);
  }
  return starts;
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
