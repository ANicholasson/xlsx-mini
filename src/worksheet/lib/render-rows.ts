import { XlsxMiniError, atPath, describeValue } from "../../errors/index.js";
import type { StringTable } from "../../shared-strings/index.js";
import type { CellStyle } from "../../styles/index.js";
import { type TextSink, escapeXml, needsSpacePreserve } from "../../xml/index.js";
import { columnLetter } from "./cell-ref.js";
import type { Cell, CellObject, Coerce, Row, TemporalLike } from "./model.js";
import { serialFromCivil, serialFromEpochMillis } from "./serial.js";

const MAX_ROWS = 1_048_576;
const MAX_COLUMNS = 16_384;
const MAX_TEXT_LENGTH = 32_767;
/** Built-in number formats a date-like value falls back to. */
const GENERAL = 0;
const DATE = 14;
const DATE_TIME = 22;

/** Everything a worksheet needs from its surroundings to render. */
export interface WorksheetContext {
  readonly strings: StringTable;
  /** What the worksheet needs from the styles feature: a cell-format index for a style, and validation. */
  readonly styles: {
    indexFor(
      cell: CellStyle | undefined,
      column: CellStyle | undefined,
      dateDefault: number,
    ): number;
    validate(style: CellStyle): void;
  };
  readonly coerce: Coerce | undefined;
  /** Zero-based position of the sheet in the workbook, for error paths. */
  readonly sheetIndex: number;
  /** Write text into the cells rather than through the shared-strings table. */
  readonly inlineStrings?: boolean;
}

/** What was written to the sink: the used range of cells, and whether any row element was emitted. */
export interface SheetData {
  /** The used range such as `A1:C4`, or an empty string when no cell was written. */
  readonly range: string;
  /** True when at least one `<row>` was written, including rows that only carry a height. */
  readonly hasRows: boolean;
}

/** Takes rows one at a time; `end` writes the trailing heights and reports what was written. */
export interface SheetDataWriter {
  row(row: Row): void;
  end(): SheetData;
}

/** `renderRow` wrote no row element. */
const NOTHING = -2;
/** `renderRow` wrote a row with a height and no cells. */
const HEIGHT_ONLY = -1;

/**
 * Renders every row with at least one cell, or a height, into the sink; positions survive
 * because rows and cells carry `r`. The loop keeps its counters in locals: this is the byte
 * path, and the stream's row-at-a-time writer below shares `renderRow` with it.
 */
export function renderSheetData(
  rows: Iterable<Row>,
  columnStyles: readonly (CellStyle | undefined)[],
  context: WorksheetContext,
  sink: TextSink,
  rowHeights?: Readonly<Record<number, number>>,
): SheetData {
  let rowIndex = -1;
  let lastRowWithCells = 0;
  let maxColumn = -1;
  let hasRows = false;
  for (const row of rows) {
    rowIndex++;
    const written = renderRow(row, rowIndex, columnStyles, context, sink, rowHeights);
    if (written === NOTHING) continue;
    hasRows = true;
    if (written === HEIGHT_ONLY) continue;
    lastRowWithCells = rowIndex + 1;
    if (written > maxColumn) maxColumn = written;
  }
  return endSheetData(rowIndex, lastRowWithCells, maxColumn, hasRows, sink, rowHeights);
}

/** The row loop as an object, so a stream can feed it one row per pull. */
export function createSheetDataWriter(
  columnStyles: readonly (CellStyle | undefined)[],
  context: WorksheetContext,
  sink: TextSink,
  rowHeights?: Readonly<Record<number, number>>,
): SheetDataWriter {
  let rowIndex = -1;
  let lastRowWithCells = 0;
  let maxColumn = -1;
  let hasRows = false;
  return {
    row(row) {
      rowIndex++;
      const written = renderRow(row, rowIndex, columnStyles, context, sink, rowHeights);
      if (written === NOTHING) return;
      hasRows = true;
      if (written === HEIGHT_ONLY) return;
      lastRowWithCells = rowIndex + 1;
      if (written > maxColumn) maxColumn = written;
    },
    end: () => endSheetData(rowIndex, lastRowWithCells, maxColumn, hasRows, sink, rowHeights),
  };
}

/**
 * Writes one row and returns the highest column written, `HEIGHT_ONLY` for a row element
 * without cells, or `NOTHING`.
 */
function renderRow(
  row: Row,
  rowIndex: number,
  columnStyles: readonly (CellStyle | undefined)[],
  context: WorksheetContext,
  sink: TextSink,
  rowHeights: Readonly<Record<number, number>> | undefined,
): number {
  if (rowIndex >= MAX_ROWS) {
    throw new XlsxMiniError(
      "TOO_MANY_ROWS",
      rowPath(context, rowIndex),
      `Excel allows at most ${MAX_ROWS} rows`,
    );
  }
  if (row.length > MAX_COLUMNS) {
    throw new XlsxMiniError(
      "TOO_MANY_COLUMNS",
      rowPath(context, rowIndex),
      `Excel allows at most ${MAX_COLUMNS} columns`,
    );
  }
  let cells = "";
  let maxColumn = -1;
  for (let column = 0; column < row.length; column++) {
    const cell = row[column];
    if (cell === null || cell === undefined) continue;
    const rendered = renderCell(cell, column, columnStyles[column], rowIndex, context);
    if (rendered.length === 0) continue;
    cells += rendered;
    maxColumn = column;
  }
  const rowNumber = rowIndex + 1;
  const height = rowHeights?.[rowNumber];
  if (cells.length === 0) {
    if (height === undefined) return NOTHING;
    sink.write(`<row r="${rowNumber}"${heightAttributes(height)}/>`);
    return HEIGHT_ONLY;
  }
  sink.write(
    `<row r="${rowNumber}"${height === undefined ? "" : heightAttributes(height)}>${cells}</row>`,
  );
  return maxColumn;
}

function endSheetData(
  rowIndex: number,
  lastRowWithCells: number,
  maxColumn: number,
  hasRows: boolean,
  sink: TextSink,
  rowHeights: Readonly<Record<number, number>> | undefined,
): SheetData {
  if (rowHeights !== undefined && renderTrailingHeights(rowHeights, rowIndex + 1, sink)) {
    hasRows = true;
  }
  const range = lastRowWithCells === 0 ? "" : `A1:${columnLetter(maxColumn)}${lastRowWithCells}`;
  return { range, hasRows };
}

function heightAttributes(height: number): string {
  return ` ht="${height}" customHeight="1"`;
}

/**
 * Heights of rows past the last row given are written as empty rows, in order, so they still
 * apply. Returns true when any row was written.
 */
function renderTrailingHeights(
  rowHeights: Readonly<Record<number, number>>,
  rowsGiven: number,
  sink: TextSink,
): boolean {
  const trailing: number[] = [];
  for (const key of Object.keys(rowHeights)) {
    const row = Number(key);
    if (row > rowsGiven) trailing.push(row);
  }
  if (trailing.length === 0) return false;
  trailing.sort((a, b) => a - b);
  for (const row of trailing) sink.write(`<row r="${row}"${heightAttributes(rowHeights[row]!)}/>`);
  return true;
}

function renderCell(
  cell: Cell,
  column: number,
  columnStyle: CellStyle | undefined,
  rowIndex: number,
  context: WorksheetContext,
): string {
  let value: unknown = cell;
  let cellStyle: CellStyle | undefined;
  if (isWrapper(value)) {
    cellStyle = value.style;
    value = value.value;
    if (isWrapper(value)) {
      throw new XlsxMiniError(
        "UNSUPPORTED_VALUE",
        cellPath(context, rowIndex, column),
        "A cell object cannot wrap another cell object",
      );
    }
  }
  const reference = columnLetter(column) + (rowIndex + 1);
  if (value === null || value === undefined) {
    // A cell object with a style and no value is a styled empty cell: a border or fill can run
    // across a blank. A bare null stays nothing, even in a styled column.
    if (cellStyle === undefined) return "";
    const style = styleAttribute(cellStyle, columnStyle, GENERAL, column, rowIndex, context);
    return style.length === 0 ? "" : `<c r="${reference}"${style}/>`;
  }
  const style =
    cellStyle === undefined && columnStyle === undefined
      ? ""
      : styleAttribute(cellStyle, columnStyle, GENERAL, column, rowIndex, context);
  return renderValue(
    value,
    reference,
    style,
    cellStyle,
    columnStyle,
    column,
    rowIndex,
    context,
    false,
  );
}

function renderValue(
  value: unknown,
  reference: string,
  style: string,
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  column: number,
  rowIndex: number,
  context: WorksheetContext,
  coerced: boolean,
): string {
  if (value === null || value === undefined) return "";
  switch (typeof value) {
    case "string": {
      if (value.length > MAX_TEXT_LENGTH) {
        throw new XlsxMiniError(
          "CELL_TEXT_TOO_LONG",
          cellPath(context, rowIndex, column),
          `Excel allows at most ${MAX_TEXT_LENGTH} characters in a cell`,
        );
      }
      if (context.inlineStrings === true) {
        const space = needsSpacePreserve(value) ? ' xml:space="preserve"' : "";
        return `<c r="${reference}"${style} t="inlineStr"><is><t${space}>${escapeXml(value)}</t></is></c>`;
      }
      return `<c r="${reference}"${style} t="s"><v>${context.strings.intern(value)}</v></c>`;
    }
    case "number": {
      if (!Number.isFinite(value)) {
        throw new XlsxMiniError(
          "INVALID_NUMBER",
          cellPath(context, rowIndex, column),
          `${value} cannot be written to a cell`,
        );
      }
      return `<c r="${reference}"${style}><v>${value}</v></c>`;
    }
    case "boolean":
      return `<c r="${reference}"${style} t="b"><v>${value ? 1 : 0}</v></c>`;
    default:
      break;
  }

  // Everything else: Date, Temporal-like, registered, or unsupported. The coerce hook sees it once.
  if (!coerced && context.coerce !== undefined) {
    const next = context.coerce(value as Date | TemporalLike, cellPath(context, rowIndex, column));
    return renderValue(
      next,
      reference,
      style,
      cellStyle,
      columnStyle,
      column,
      rowIndex,
      context,
      true,
    );
  }
  if (value instanceof Date) {
    const millis = value.getTime();
    if (Number.isNaN(millis)) {
      throw new XlsxMiniError(
        "INVALID_DATE",
        cellPath(context, rowIndex, column),
        "Invalid Date cannot be written to a cell",
      );
    }
    return dateCell(
      reference,
      serialFromEpochMillis(millis),
      DATE_TIME,
      cellStyle,
      columnStyle,
      column,
      rowIndex,
      context,
    );
  }
  if (typeof value === "object") {
    const tag = (value as { readonly [Symbol.toStringTag]?: unknown })[Symbol.toStringTag];
    switch (tag) {
      case "Temporal.PlainDate": {
        const date = value as { year: number; month: number; day: number };
        const serial = serialFromCivil(date.year, date.month, date.day, 0, 0, 0, 0);
        return dateCell(reference, serial, DATE, cellStyle, columnStyle, column, rowIndex, context);
      }
      case "Temporal.PlainDateTime":
      case "Temporal.ZonedDateTime": {
        const time = value as {
          year: number;
          month: number;
          day: number;
          hour: number;
          minute: number;
          second: number;
          millisecond: number;
          microsecond: number;
          nanosecond: number;
        };
        const millisecond =
          time.millisecond + time.microsecond / 1000 + time.nanosecond / 1_000_000;
        const serial = serialFromCivil(
          time.year,
          time.month,
          time.day,
          time.hour,
          time.minute,
          time.second,
          millisecond,
        );
        return dateCell(
          reference,
          serial,
          DATE_TIME,
          cellStyle,
          columnStyle,
          column,
          rowIndex,
          context,
        );
      }
      case "Temporal.Instant": {
        const serial = serialFromEpochMillis(
          (value as { epochMilliseconds: number }).epochMilliseconds,
        );
        return dateCell(
          reference,
          serial,
          DATE_TIME,
          cellStyle,
          columnStyle,
          column,
          rowIndex,
          context,
        );
      }
      default:
        break;
    }
  }
  throw new XlsxMiniError(
    "UNSUPPORTED_VALUE",
    cellPath(context, rowIndex, column),
    `A value of type ${describe(value)} cannot be written to a cell; register the type and convert it in coerce`,
  );
}

function dateCell(
  reference: string,
  serial: number,
  dateDefault: number,
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  column: number,
  rowIndex: number,
  context: WorksheetContext,
): string {
  if (Number.isNaN(serial)) {
    throw new XlsxMiniError(
      "DATE_OUT_OF_RANGE",
      cellPath(context, rowIndex, column),
      "Excel dates run from 1900-01-01 to 9999-12-31",
    );
  }
  const style = styleAttribute(cellStyle, columnStyle, dateDefault, column, rowIndex, context);
  return `<c r="${reference}"${style}><v>${serial}</v></c>`;
}

/** The ` s="N"` attribute for a styled cell, or an empty string for the default format. */
function styleAttribute(
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  dateDefault: number,
  column: number,
  rowIndex: number,
  context: WorksheetContext,
): string {
  let index: number;
  try {
    index = context.styles.indexFor(cellStyle, columnStyle, dateDefault);
  } catch (error) {
    if (error instanceof XlsxMiniError) {
      throw atPath(error, `${cellPath(context, rowIndex, column)}.style`);
    }
    throw error;
  }
  return index === 0 ? "" : ` s="${index}"`;
}

/** A plain object with a `value` property is a cell object; class instances never are. */
function isWrapper(value: unknown): value is CellObject {
  if (typeof value !== "object" || value === null) return false;
  const prototype = Object.getPrototypeOf(value) as unknown;
  return (prototype === Object.prototype || prototype === null) && "value" in value;
}

function describe(value: unknown): string {
  return typeof value !== "object" || value === null ? typeof value : describeValue(value);
}

function cellPath(context: WorksheetContext, rowIndex: number, column: number): string {
  return `sheets[${context.sheetIndex}].rows[${rowIndex}][${column}]`;
}

function rowPath(context: WorksheetContext, rowIndex: number): string {
  return `sheets[${context.sheetIndex}].rows[${rowIndex}]`;
}
