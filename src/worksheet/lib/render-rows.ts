import { XlsxMiniError, atPath, describeValue } from "../../errors/index.js";
import type { StringTable } from "../../shared-strings/index.js";
import type { CellStyle } from "../../styles/index.js";
import { type TextSink, escapeXml, needsSpacePreserve } from "../../xml/index.js";
import { IMAGE_INSET, type MediaTable, type SheetDrawing } from "../../drawing/index.js";
import type { AutoWidths } from "./auto-width.js";
import { MAX_COLUMNS, MAX_ROWS, columnLetter } from "./cell-ref.js";
import type { Hyperlinks } from "./hyperlinks.js";
import type {
  Cell,
  CellObject,
  Coerce,
  FormulaCell,
  ImageCell,
  LinkCell,
  Row,
  SheetLinkCell,
  TemporalLike,
} from "./model.js";
import { serialFromCivil, serialFromEpochMillis } from "./serial.js";

const MAX_TEXT_LENGTH = 32_767;
const MAX_FORMULA_LENGTH = 8192;
const MAX_ROW_HEIGHT = 409;
/** The height of a row nobody sized, in points; a picture never makes a row smaller than that. */
const DEFAULT_ROW_HEIGHT = 15;
/** Excel's 96 dpi against 72 points per inch. */
const POINTS_PER_PIXEL = 0.75;
/** The tallest picture a row can make room for: the row limit in pixels, less the inset. */
const MAX_IMAGE_ROW_PIXELS = Math.floor(MAX_ROW_HEIGHT / POINTS_PER_PIXEL) - 2 * IMAGE_INSET;
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
  /** Every sheet's name in tab order, so a link to another sheet can be checked. */
  readonly sheetNames: readonly string[];
  /** The workbook's images, shared by every sheet so a repeated image is written once. */
  readonly media: MediaTable;
  /** Write text into the cells rather than through the shared-strings table. */
  readonly inlineStrings?: boolean;
}

/** What the rows leave behind for the parts written after them. */
export interface RowRecords {
  readonly links: Hyperlinks;
  /** Set once a formula has been written, so the workbook can ask Excel to calculate on open. */
  hasFormulas: boolean;
  /** The measurer for the columns that size themselves, or undefined when none does. */
  readonly widths: AutoWidths | undefined;
  /** The sheet's pictures, anchored as their cells are written. */
  readonly drawing: SheetDrawing;
  /** The row the picture records below belong to; a row without pictures never touches them. */
  imageRow: number;
  /** The tallest picture in that row, in pixels, and the column holding it. */
  rowImagePixels: number;
  rowTallestColumn: number;
  /** The rightmost column holding a picture in that row. */
  rowImageColumn: number;
  /** The widest picture per column, in pixels, for the column widths. */
  readonly imagePixels: number[];
}

/** The worksheet context with the sheet's records, so the per-cell calls carry one argument for both. */
export interface RowContext extends WorksheetContext {
  readonly records: RowRecords;
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
  context: RowContext,
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
  context: RowContext,
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
  context: RowContext,
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
  let height = rowHeights?.[rowNumber];
  const records = context.records;
  if (records.imageRow === rowIndex) {
    // A picture occupies its cell whether or not a cell element is written beneath it.
    if (records.rowImageColumn > maxColumn) maxColumn = records.rowImageColumn;
    // The caller's height wins; otherwise the row makes room for its tallest picture.
    if (height === undefined) {
      height = imageRowHeight(records.rowImagePixels);
      if (height === undefined) {
        throw new XlsxMiniError(
          "INVALID_IMAGE",
          cellPath(context, rowIndex, records.rowTallestColumn),
          `A picture ${records.rowImagePixels} pixels tall does not fit a row, which Excel caps at ${MAX_ROW_HEIGHT} points (${MAX_IMAGE_ROW_PIXELS} pixels); give it a height, or the row one`,
        );
      }
    }
  }
  if (cells.length === 0) {
    if (height === undefined) return NOTHING;
    sink.write(`<row r="${rowNumber}"${heightAttributes(height)}/>`);
    return maxColumn === -1 ? HEIGHT_ONLY : maxColumn;
  }
  sink.write(
    `<row r="${rowNumber}"${height === undefined ? "" : heightAttributes(height)}>${cells}</row>`,
  );
  return maxColumn;
}

/** Writes the heights of rows past the last one given and reports the used range. */
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

/** The points a row needs for a picture this many pixels tall, or undefined when no row can hold it. */
function imageRowHeight(pixels: number): number | undefined {
  if (pixels > MAX_IMAGE_ROW_PIXELS) return undefined;
  const points = Math.round((pixels + 2 * IMAGE_INSET) * POINTS_PER_PIXEL * 100) / 100;
  return points < DEFAULT_ROW_HEIGHT ? DEFAULT_ROW_HEIGHT : points;
}

/** The row attributes for a custom height in points. */
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

/** One `<c>` element for a cell, or an empty string when nothing is written; link cells are also recorded. */
function renderCell(
  cell: Cell,
  column: number,
  columnStyle: CellStyle | undefined,
  rowIndex: number,
  context: RowContext,
): string {
  const records = context.records;
  let value: unknown = cell;
  let cellStyle: CellStyle | undefined;
  const reference = columnLetter(column) + (rowIndex + 1);
  if (isWrapper(value)) {
    cellStyle = value.style;
    if ("image" in value) {
      return renderImage(value, reference, cellStyle, columnStyle, column, rowIndex, context);
    }
    if ("formula" in value) {
      if ("href" in value || "cell" in value || "sheet" in value) {
        throw new XlsxMiniError(
          "INVALID_FORMULA",
          cellPath(context, rowIndex, column),
          "A formula cell cannot also be a link",
        );
      }
      records.hasFormulas = true;
      return renderFormula(value, reference, cellStyle, columnStyle, column, rowIndex, context);
    }
    if ("href" in value || "cell" in value || "sheet" in value) {
      records.links.add(
        reference,
        value as LinkCell | SheetLinkCell,
        cellPath(context, rowIndex, column),
      );
    }
    value = value.value;
    if (isWrapper(value)) {
      throw new XlsxMiniError(
        "UNSUPPORTED_VALUE",
        cellPath(context, rowIndex, column),
        "A cell object cannot wrap another cell object",
      );
    }
  }
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

/** A cell holding a plain value, after the coerce hook has had its say. */
function renderValue(
  value: unknown,
  reference: string,
  style: string,
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  column: number,
  rowIndex: number,
  context: RowContext,
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
      const widths = context.records.widths;
      const measure = widths !== undefined && widths.wants(column, rowIndex);
      if (context.inlineStrings === true) {
        if (measure) widths.text(column, value, cellStyle, columnStyle);
        const space = needsSpacePreserve(value) ? ' xml:space="preserve"' : "";
        return `<c r="${reference}"${style} t="inlineStr"><is><t${space}>${escapeXml(value)}</t></is></c>`;
      }
      const index = context.strings.intern(value);
      if (measure) widths.sharedText(column, index, value, cellStyle, columnStyle);
      return `<c r="${reference}"${style} t="s"><v>${index}</v></c>`;
    }
    case "number": {
      if (!Number.isFinite(value)) {
        throw new XlsxMiniError(
          "INVALID_NUMBER",
          cellPath(context, rowIndex, column),
          `${value} cannot be written to a cell`,
        );
      }
      const widths = context.records.widths;
      if (widths !== undefined && widths.wants(column, rowIndex)) {
        widths.number(column, value, GENERAL, cellStyle, columnStyle);
      }
      return `<c r="${reference}"${style}><v>${value}</v></c>`;
    }
    case "boolean": {
      const widths = context.records.widths;
      if (widths !== undefined && widths.wants(column, rowIndex)) {
        widths.text(column, value ? "TRUE" : "FALSE", cellStyle, columnStyle);
      }
      return `<c r="${reference}"${style} t="b"><v>${value ? 1 : 0}</v></c>`;
    }
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

/**
 * A formula cell: the formula escaped inside `<f>`, never evaluated, and the cached value, when
 * given, typed as the cell's own value would be. Text is `str`, since a result is never shared.
 */
function renderFormula(
  cell: FormulaCell,
  reference: string,
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  column: number,
  rowIndex: number,
  context: RowContext,
): string {
  const path = cellPath(context, rowIndex, column);
  const formula = cell.formula;
  if (typeof formula !== "string" || formula.length === 0 || formula.length > MAX_FORMULA_LENGTH) {
    throw new XlsxMiniError(
      "INVALID_FORMULA",
      `${path}.formula`,
      `A formula is 1 to ${MAX_FORMULA_LENGTH} characters, not ${describeValue(formula)}`,
    );
  }
  if (formula.startsWith("=")) {
    throw new XlsxMiniError(
      "INVALID_FORMULA",
      `${path}.formula`,
      "A formula is written without the leading =",
    );
  }
  const style =
    cellStyle === undefined && columnStyle === undefined
      ? ""
      : styleAttribute(cellStyle, columnStyle, GENERAL, column, rowIndex, context);
  const open = `<c r="${reference}"${style}`;
  const f = `<f>${escapeXml(formula)}</f>`;
  const value = cell.value;
  if (value === null || value === undefined) return `${open}>${f}</c>`;
  const widths = context.records.widths;
  const measure = widths !== undefined && widths.wants(column, rowIndex);
  switch (typeof value) {
    case "string":
      if (value.length > MAX_TEXT_LENGTH) {
        throw new XlsxMiniError(
          "CELL_TEXT_TOO_LONG",
          `${path}.value`,
          `Excel allows at most ${MAX_TEXT_LENGTH} characters in a cell`,
        );
      }
      if (measure) widths.text(column, value, cellStyle, columnStyle);
      return `${open} t="str">${f}<v>${escapeXml(value)}</v></c>`;
    case "number":
      if (!Number.isFinite(value)) {
        throw new XlsxMiniError(
          "INVALID_NUMBER",
          `${path}.value`,
          `${value} cannot be written to a cell`,
        );
      }
      if (measure) widths.number(column, value, GENERAL, cellStyle, columnStyle);
      return `${open}>${f}<v>${value}</v></c>`;
    case "boolean":
      if (measure) widths.text(column, value ? "TRUE" : "FALSE", cellStyle, columnStyle);
      return `${open} t="b">${f}<v>${value ? 1 : 0}</v></c>`;
    default:
      throw new XlsxMiniError(
        "INVALID_FORMULA",
        `${path}.value`,
        `A cached formula value is text, a number or a boolean, not ${describe(value)}`,
      );
  }
}

/** A numeric cell holding a date serial, styled with a date format unless the style sets one. */
function dateCell(
  reference: string,
  serial: number,
  dateDefault: number,
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  column: number,
  rowIndex: number,
  context: RowContext,
): string {
  if (Number.isNaN(serial)) {
    throw new XlsxMiniError(
      "DATE_OUT_OF_RANGE",
      cellPath(context, rowIndex, column),
      "Excel dates run from 1900-01-01 to 9999-12-31",
    );
  }
  const widths = context.records.widths;
  if (widths !== undefined && widths.wants(column, rowIndex)) {
    widths.number(column, serial, dateDefault, cellStyle, columnStyle);
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

/**
 * An image cell: the picture is registered once per distinct image and anchored to this cell; the
 * row and the column learn its size. The cell beneath is written only when it carries a style.
 */
function renderImage(
  cell: ImageCell,
  reference: string,
  cellStyle: CellStyle | undefined,
  columnStyle: CellStyle | undefined,
  column: number,
  rowIndex: number,
  context: RowContext,
): string {
  const path = cellPath(context, rowIndex, column);
  if ("value" in cell || "formula" in cell || "href" in cell || "cell" in cell || "sheet" in cell) {
    throw new XlsxMiniError(
      "INVALID_IMAGE",
      path,
      "An image cell holds the image, its width and height and a style, not a value, formula or link",
    );
  }
  const records = context.records;
  const media = context.media.add(cell.image, path);
  const size = records.drawing.add(column, rowIndex, media, cell.width, cell.height, path);
  if (records.imageRow !== rowIndex) {
    records.imageRow = rowIndex;
    records.rowImagePixels = 0;
  }
  if (size.height > records.rowImagePixels) {
    records.rowImagePixels = size.height;
    records.rowTallestColumn = column;
  }
  // Cells are visited left to right, so the latest picture is the rightmost.
  records.rowImageColumn = column;
  if (size.width > (records.imagePixels[column] ?? 0)) records.imagePixels[column] = size.width;
  if (cellStyle === undefined) return "";
  const style = styleAttribute(cellStyle, columnStyle, GENERAL, column, rowIndex, context);
  return style.length === 0 ? "" : `<c r="${reference}"${style}/>`;
}

/** A plain object with a value, formula, link or image property is a cell object; class instances never are. */
function isWrapper(
  value: unknown,
): value is CellObject | FormulaCell | LinkCell | SheetLinkCell | ImageCell {
  if (typeof value !== "object" || value === null) return false;
  const prototype = Object.getPrototypeOf(value) as unknown;
  return (
    (prototype === Object.prototype || prototype === null) &&
    ("value" in value ||
      "formula" in value ||
      "href" in value ||
      "cell" in value ||
      "sheet" in value ||
      "image" in value)
  );
}

/** A value's type for an error message: the typeof for primitives, the tag or constructor for objects. */
function describe(value: unknown): string {
  return typeof value !== "object" || value === null ? typeof value : describeValue(value);
}

/** The error path of a cell. */
function cellPath(context: WorksheetContext, rowIndex: number, column: number): string {
  return `sheets[${context.sheetIndex}].rows[${rowIndex}][${column}]`;
}

/** The error path of a row. */
function rowPath(context: WorksheetContext, rowIndex: number): string {
  return `sheets[${context.sheetIndex}].rows[${rowIndex}]`;
}
