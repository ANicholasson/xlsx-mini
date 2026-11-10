import type { CellStyle } from "../../styles/index.js";
import type { Cell, CellValue, Column, Row, Sheet } from "./model.js";

/** The keys of an object whose values can go straight into a cell. */
type CellKeys<T> = { [K in keyof T]-?: T[K] extends CellValue ? K : never }[keyof T] & string;

interface ColumnBase {
  /** Header text; the key when omitted. */
  readonly header?: string;
  /** Width in Excel's character units, or `"auto"` to fit the column to its content. */
  readonly width?: number | "auto";
  /** Default style for the column's data cells. */
  readonly style?: CellStyle;
}

/** A column read straight from a property whose type a cell can hold. */
interface ValueColumn<T> extends ColumnBase {
  readonly key: CellKeys<T>;
  readonly value?: undefined;
}

/** A column computed from the whole object; the key only names it. */
interface DerivedColumn<T> extends ColumnBase {
  readonly key: keyof T & string;
  readonly value: (item: T) => CellValue;
}

export type ObjectColumn<T> = ValueColumn<T> | DerivedColumn<T>;

export interface SheetFromObjectsOptions<T> {
  readonly name: string;
  /** Columns in order. When omitted, the first object's own keys become the columns, keys as headers. */
  readonly columns?: readonly ObjectColumn<T>[];
  /** Style of the header row, or `false` for no header row. Default: bold. */
  readonly header?: CellStyle | false;
  /** Keep the header row visible while scrolling. Default: on when there is a header row. */
  readonly freeze?: boolean;
  /** Filter dropdowns on the header row. Default: on when there is a header row. */
  readonly autoFilter?: boolean;
  /** Fit every column without a numeric width to its content. Default: off. */
  readonly autoWidth?: boolean;
}

const BOLD_HEADER: CellStyle = { bold: true };

/**
 * Turns objects into a sheet: one row per object, one column per key. Rows are produced lazily, so
 * a generator or cursor is read once, when the sheet is written.
 *
 * @example
 * const sheet = sheetFromObjects(orders, {
 *   name: "Orders",
 *   columns: [
 *     { key: "id", header: "Order" },
 *     { key: "total", header: "Total", style: { numberFormat: '#,##0.00 "kr"' } },
 *     { key: "tags", header: "Tags", value: (order) => order.tags.join(", ") },
 *   ],
 * });
 */
export function sheetFromObjects<T extends object>(
  data: Iterable<T>,
  options: SheetFromObjectsOptions<T>,
): Sheet {
  const headerStyle = options.header === undefined ? BOLD_HEADER : options.header;
  const hasHeader = headerStyle !== false;
  const freeze = options.freeze ?? hasHeader;
  const autoFilter = options.autoFilter ?? hasHeader;
  const columns = options.columns;
  const rows: Iterable<Row> = {
    [Symbol.iterator]: () =>
      columns === undefined
        ? inferredRows(data, headerStyle)
        : explicitRows(data, columns, headerStyle),
  };
  let sheet: Sheet = freeze
    ? { name: options.name, rows, freeze: { rows: 1 }, autoFilter }
    : { name: options.name, rows, autoFilter };
  if (options.autoWidth === true) sheet = { ...sheet, autoWidth: true };
  return columns === undefined ? sheet : { ...sheet, columns: columns.map(toColumn) };
}

function toColumn(column: ColumnBase): Column {
  const result: { width?: number | "auto"; style?: CellStyle } = {};
  if (column.width !== undefined) result.width = column.width;
  if (column.style !== undefined) result.style = column.style;
  return result;
}

function* explicitRows<T extends object>(
  data: Iterable<T>,
  columns: readonly ObjectColumn<T>[],
  headerStyle: CellStyle | false,
): Generator<Row> {
  if (headerStyle !== false)
    yield columns.map((column) => ({ value: column.header ?? column.key, style: headerStyle }));
  for (const item of data) {
    const row: Cell[] = [];
    for (const column of columns) {
      row.push(column.value === undefined ? (item[column.key] as CellValue) : column.value(item));
    }
    yield row;
  }
}

function* inferredRows<T extends object>(
  data: Iterable<T>,
  headerStyle: CellStyle | false,
): Generator<Row> {
  const iterator = data[Symbol.iterator]();
  const first = iterator.next();
  if (first.done === true) return;
  const keys = Object.keys(first.value) as (keyof T & string)[];
  if (headerStyle !== false) yield keys.map((key) => ({ value: key, style: headerStyle }));
  yield keys.map((key) => first.value[key] as CellValue);
  for (let next = iterator.next(); next.done !== true; next = iterator.next()) {
    const item = next.value;
    yield keys.map((key) => item[key] as CellValue);
  }
}
