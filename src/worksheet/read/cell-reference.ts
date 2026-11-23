/**
 * Cell references on the read path. The write side's `parseReference` takes ranges and whole
 * columns from a caller and matches patterns; the read side scans one cell in place, once per cell.
 * The bounds and the column letters are the write side's.
 */
import { parseIndex } from "../../xml/read.js";
import { MAX_COLUMNS, columnLetter } from "../lib/cell-ref.js";

const A = 65;
const Z = 90;
const ZERO = 48;

/**
 * The column, from 1, that the reference `buffer[start, end)` names in row `row`, or 0 when it is
 * not one to three capital letters and then exactly the digits of `row`, without a leading zero,
 * or names a column past the sheet's last.
 */
export function columnOfReference(buffer: string, start: number, end: number, row: number): number {
  let column = 0;
  let at = start;
  for (let code = buffer.charCodeAt(at); code >= A && code <= Z && at - start < 3;) {
    column = column * 26 + code - (A - 1);
    code = buffer.charCodeAt(++at);
  }
  return column > MAX_COLUMNS ||
    buffer.charCodeAt(at) === ZERO ||
    parseIndex(buffer, at, end) !== row
    ? 0
    : column;
}

/** The reference of a cell for an error: column 3 of row 41 is `C41`. */
export function referenceOf(column: number, row: number): string {
  return columnLetter(column - 1) + row;
}
