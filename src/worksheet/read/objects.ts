/**
 * Objects by declared columns: the mirror of `sheetFromObjects`. Keys come from the caller and
 * never from the file; text from the file is only looked up among the declared headers.
 */
import { XlsxReadError } from "../../errors/read.js";
import type { ReadRow, ReadValue } from "./model.js";

/**
 * One declared column: the key its value is given under, and the header that finds it in the
 * header row. A column of the writer's `sheetFromObjects` is one as it is, so one list serves the
 * export and the import; what the writer adds to a column, such as a width, a style or a `value`
 * function, is not read.
 *
 * @example
 * const columns = [{ key: "id", header: "Order" }, { key: "total", header: "Total" }] as const;
 *
 * @experimental
 */
export interface ReadColumn<Key extends string> {
  readonly key: Key;
  /** The text of the column's header cell, matched exactly. The key when omitted. */
  readonly header?: string;
}

/**
 * One row after the header row, as `objectsFromRows` yields it.
 *
 * @example
 * for await (const { number, values } of objectsFromRows(book.rows(0), columns)) {
 *   if (typeof values.total !== "number") console.log(`Row ${number}: Total must be a number`);
 * }
 *
 * @experimental
 */
export interface ReadObject<Key extends string> {
  /** The sheet's row number, from 1. */
  readonly number: number;
  /**
   * Every declared key, in the order of the columns, with the value of its column in this row, or
   * `null` where the row ends before that column. A new object for every row, with no prototype: a
   * key named `__proto__` or `constructor` is a value like any other, and the object has no
   * `toString` or `hasOwnProperty`, so `String(values)` throws. `JSON.stringify` and a spread copy
   * its values.
   */
  readonly values: { readonly [K in Key]: ReadValue };
}

const MISSING = "the header row lacks a declared column";

function headerOf({ key, header }: ReadColumn<string>): string {
  return header ?? key;
}

function missing(column: ReadColumn<string> | undefined): never {
  throw new XlsxReadError("MISSING_HEADER", MISSING, {
    header: column === undefined ? undefined : headerOf(column),
  });
}

/** A declared key and the position of its value in a row, from 0. */
type Field = readonly [key: string, position: number];

/**
 * Where each key's value stands in a row, read from the header row once. Only text is declared, so
 * a cell that holds anything else finds no column. Columns take the cells that hold their header
 * from the left, so the k-th column declared with a header takes the k-th cell that holds it, and a
 * list with one header twice reads back what the writer wrote from it. Of two columns with one key,
 * the first gives the value.
 */
function fieldsOf(header: readonly ReadValue[], columns: readonly ReadColumn<string>[]): Field[] {
  const found = new Map<ReadValue | undefined, number[]>();
  for (let position = header.length - 1; position >= 0; position--) {
    const cell = header[position];
    const positions = found.get(cell);
    if (positions === undefined) found.set(cell, [position]);
    else positions.push(position);
  }
  const keys = new Set<string>();
  const fields: Field[] = [];
  for (const column of columns) {
    const position = found.get(headerOf(column))?.pop();
    if (position === undefined) missing(column);
    if (keys.has(column.key)) continue;
    keys.add(column.key);
    fields.push([column.key, position]);
  }
  return fields;
}

/**
 * Typed with the keys as `string`: the types cannot follow a key from the list into the object, and
 * TypeScript takes a `ReadObject<string>` for a `ReadObject` of any keys.
 */
async function* objects(
  rows: AsyncIterable<ReadRow>,
  columns: readonly ReadColumn<string>[],
): AsyncGenerator<ReadObject<string>, void, undefined> {
  let fields: readonly Field[] | undefined;
  for await (const { number, cells } of rows) {
    if (fields === undefined) {
      fields = fieldsOf(cells, columns);
      continue;
    }
    const values: Record<string, ReadValue> = { __proto__: null };
    for (const [key, position] of fields) values[key] = cells[position] ?? null;
    yield { number, values };
  }
  if (fields === undefined) missing(columns[0]);
}

/**
 * Reads rows as objects by declared columns, the mirror of `sheetFromObjects`. The first row the
 * rows yield is the header row, whatever its number; to skip a title above it, hand over the rows
 * from the header row on. Every later row yields one object. Keys come from the caller and never
 * from the file: a header the caller did not declare is ignored.
 *
 * A header cell matches a column when it holds text equal to the column's header, or to its key
 * when it has none. A number, a boolean, a date or an empty cell matches no column, since the
 * writer writes every header as text and the text an application shows for a number depends on its
 * format; `rows` reads empty text as `null`, so a column with an empty header is never found in a
 * sheet. Of two cells with one header the first is the column's, and a second column with that
 * header takes the second cell. Of two columns with one key, the first gives the value.
 *
 * Throws `MISSING_HEADER` for the first declared column the header row lacks, with that column's
 * header in `header`, and when the rows yield none at all, with the first column's. It is thrown
 * by the first step of an iteration, not by the call: nothing is known of the header before it is
 * read. `header` is the caller's own text, and `sheet` is not set, since the rows do not say which
 * sheet they are. An error of the rows passes through as it is.
 *
 * Nothing is read until the result is iterated, and every iteration iterates the rows again from
 * their start. Leaving the loop early ends the iteration of the rows, which for `rows` stops the
 * inflate. The header row is read once per iteration; every later row costs one object and one
 * array read per declared column, however wide the sheet. What a value must be stays with the
 * caller.
 *
 * @example
 * const columns = [{ key: "id", header: "Order" }, { key: "total", header: "Total" }] as const;
 * for await (const { number, values } of objectsFromRows(book.rows("Orders"), columns)) {
 *   if (typeof values.total !== "number") console.log(`Row ${number}: Total must be a number`);
 * }
 *
 * @experimental
 */
export function objectsFromRows<Key extends string>(
  rows: AsyncIterable<ReadRow>,
  columns: readonly ReadColumn<Key>[],
): AsyncIterable<ReadObject<Key>> {
  return { [Symbol.asyncIterator]: () => objects(rows, columns) };
}
