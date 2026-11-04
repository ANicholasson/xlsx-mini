import type { StringTable } from "../../shared-strings/index.js";
import type { Row } from "../index.js";
import { columnLetter } from "./cell-ref.js";

export interface RenderedSheetData {
  /** The `<row>` elements, or an empty string when no cell was written. */
  readonly rows: string;
  /** The `<dimension>` element, or an empty string when no cell was written. */
  readonly dimension: string;
}

/** Renders every row with at least one cell. Positions survive because rows and cells carry `r`. */
export function renderSheetData(rows: Iterable<Row>, strings: StringTable): RenderedSheetData {
  const chunks: string[] = [];
  let rowNumber = 0;
  let lastRowWithCells = 0;
  let maxColumn = -1;

  for (const row of rows) {
    rowNumber++;
    let cells = "";
    for (let column = 0; column < row.length; column++) {
      const value = row[column];
      if (value === null || value === undefined) continue;
      const reference = columnLetter(column) + rowNumber;
      cells +=
        typeof value === "string"
          ? `<c r="${reference}" t="s"><v>${strings.intern(value)}</v></c>`
          : `<c r="${reference}"><v>${value}</v></c>`;
      if (column > maxColumn) maxColumn = column;
    }
    if (cells.length === 0) continue;
    chunks.push(`<row r="${rowNumber}">${cells}</row>`);
    lastRowWithCells = rowNumber;
  }

  if (lastRowWithCells === 0) return { rows: "", dimension: "" };
  return {
    rows: chunks.join(""),
    dimension: `<dimension ref="A1:${columnLetter(maxColumn)}${lastRowWithCells}"/>`,
  };
}
