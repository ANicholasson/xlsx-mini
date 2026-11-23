/**
 * The read side of the worksheet: from the tokens of a worksheet part to rows of values. Owns the
 * rules of the sheet data: where it stands, the order of rows and cells, references, and what each
 * type of cell stores.
 */
import { type Meter, type RowMeter, XlsxReadError } from "../errors/read.js";
import { type SharedStrings, createStringItem } from "../shared-strings/read.js";
import type { XmlCursor, XmlVisitor } from "../xml/read.js";
import { MAX_COLUMNS, MAX_ROWS } from "./lib/cell-ref.js";
import { columnOfReference, referenceOf } from "./read/cell-reference.js";
import { DATE, INLINE_STRING, UNKNOWN, cellKindOf, cellValue } from "./read/cell-value.js";
import type { ReadRow, ReadValue } from "./read/model.js";

export type { ReadRow, ReadValue } from "./read/model.js";

/** What one `rows` iteration reads a worksheet with. */
export interface WorksheetContext {
  readonly strings: SharedStrings;
  readonly positions: Meter<"cellPositions">;
  readonly rowText: RowMeter<"rowCharacters">;
}

const DAMAGED = "a worksheet part is damaged or holds no sheet data";
const ROW = "a row is out of order or misnumbered";
const REFERENCE = "a cell reference is malformed or out of order";
const KIND = "a cell is of an unknown type";
const SHAPE = "a row or a cell holds what it cannot hold";
const VALUE = "a cell's value is not of its type or names no shared string";
const DATE_TEXT = "a date stored as text is not read yet";

function damaged(): never {
  throw new XlsxReadError("INVALID_WORKBOOK", DAMAGED);
}

function refuse(message: string, cell: string | undefined): never {
  throw new XlsxReadError("INVALID_CELL", message, { cell });
}

/**
 * The visitor of a worksheet part. It returns each row that holds a value as the row closes: one
 * pass, with one array and one object per such row and none per cell.
 *
 * Every value is charged before the row's array grows: `positions` the positions it adds, the gap
 * before it included, and `rowText` the length of a text, so a shared string is charged each time a
 * cell names it, and an inline string run by run as its runs are read. Nothing is sized from what
 * the file states: `dimension` and `spans` are not read.
 *
 * Refuses, never repairs: `INVALID_WORKBOOK` when the sheet data cannot be found, is stated twice,
 * or holds anything but rows, and when a `sheetData` or a `row` stands anywhere else; `INVALID_CELL`,
 * with the place in the sheet, for whatever is wrong in a row or a cell, a date stored as text
 * included. Hidden rows and columns read like visible ones.
 */
export function worksheetVisitor(context: WorksheetContext): XmlVisitor<ReadRow, void> {
  const { strings, positions, rowText } = context;
  const item = createStringItem(5);
  let pathDepth = 0;
  let skippedDepth = 0;
  let found = false;
  let previous = 0;
  let row = 0;
  let cells: ReadValue[] = [];
  let column = 0;
  let kind = 0;
  let inline = false;
  let stored: string | undefined;

  function at(position: number): string {
    return referenceOf(position, row);
  }

  function openRow(stated: number): void {
    row = stated < 0 ? previous + 1 : stated;
    if (!(row > previous && row <= MAX_ROWS)) refuse(ROW, row > 0 ? at(1) : undefined);
    previous = row;
    column = 0;
    rowText.next();
  }

  function openCell(xml: XmlCursor): void {
    const start = xml.valueStart("r");
    const next = start < 0 ? column + 1 : columnOfReference(xml.buffer, start, xml.valueEnd, row);
    if (!(next > column && next <= MAX_COLUMNS)) {
      refuse(REFERENCE, start < 0 ? at(next) : xml.attribute("r"));
    }
    column = next;
    kind = cellKindOf(xml.buffer, xml.valueStart("t"), xml.valueEnd);
    if (kind === UNKNOWN) refuse(KIND, at(column));
    stored = undefined;
  }

  function closeCell(): void {
    const value = cellValue(kind, stored, strings);
    if (value === undefined) refuse(kind === DATE ? DATE_TEXT : VALUE, at(column));
    if (value === null) return;
    positions.charge(column - cells.length);
    if (typeof value === "string" && kind !== INLINE_STRING) rowText.charge(value.length);
    while (cells.length < column - 1) cells.push(null);
    cells.push(value);
  }

  return {
    open(xml) {
      const name = xml.name;
      if (skippedDepth === 0 && pathDepth === 1 && name !== "sheetData") skippedDepth = 2;
      if (skippedDepth !== 0) {
        if (name === "sheetData" || name === "row") damaged();
        return;
      }
      switch (pathDepth) {
        case 0:
          if (name !== "worksheet") damaged();
          break;
        case 1:
          if (found) damaged();
          found = true;
          break;
        case 2:
          if (name !== "row") damaged();
          openRow(xml.index("r"));
          break;
        case 3:
          if (name !== "c") {
            if (name !== "extLst") refuse(SHAPE, at(column + 1));
            skippedDepth = 4;
            return;
          }
          openCell(xml);
          break;
        case 4:
          if (name === "f" || name === "extLst") {
            skippedDepth = 5;
            return;
          }
          inline = name === "is";
          if (
            (!inline && name !== "v") ||
            stored !== undefined ||
            inline !== (kind === INLINE_STRING)
          ) {
            refuse(SHAPE, at(column));
          }
          break;
        default:
          if (!inline || !item.open(xml)) refuse(SHAPE, at(column));
          return;
      }
      pathDepth++;
    },
    close(xml) {
      const depth = xml.depth;
      if (skippedDepth !== 0) {
        if (depth === skippedDepth) skippedDepth = 0;
        return undefined;
      }
      if (depth > pathDepth) {
        rowText.charge(item.close(xml));
        return undefined;
      }
      pathDepth--;
      if (depth === 5) stored = inline ? item.take() : xml.text();
      else if (depth === 4) closeCell();
      else if (depth === 3 && cells.length > 0) {
        const full = { number: row, cells };
        cells = [];
        return full;
      }
      return undefined;
    },
    end() {
      if (!found) damaged();
    },
  };
}
