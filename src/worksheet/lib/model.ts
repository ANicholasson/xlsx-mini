/** What a cell can hold. Grows with the cell-model work (booleans, dates, Temporal). */
export type CellValue = string | number | null | undefined;

/** One cell of a row. `null` and `undefined` write no cell but keep the column position. */
export type Cell = CellValue;

/** One row, left to right from column A. */
export type Row = readonly Cell[];

/**
 * One sheet of a workbook.
 *
 * @example
 * const sheet: Sheet = { name: "Orders", rows: [["Order", "Total"], ["A-1001", 1234.5]] };
 */
export interface Sheet {
  readonly name: string;
  /** Any iterable of rows, so a generator or database cursor can feed a sheet without an array. */
  readonly rows: Iterable<Row>;
}
