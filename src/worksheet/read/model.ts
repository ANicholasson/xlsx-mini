/**
 * The value of a cell as the file stores it:
 *
 * - text is a `string`, as stored and untrimmed, with rich text as its plain text;
 * - a number is a `number`, never `NaN` or an infinity. A date or a time is the serial the file
 *   stores, a number of days with the time of day as the fraction;
 * - a boolean is a `boolean`;
 * - an error cell is its text, such as `"#DIV/0!"`;
 * - a formula is its stored result, or `null` when the file stores none: nothing is calculated;
 * - an empty cell, and empty text, is `null`.
 *
 * In a shared string, an inline string and a formula's text result, an escape `_xHHHH_` reads as
 * the UTF-16 code unit it names, as the format's rule for the escape says, also for a character
 * XML cannot carry: `_x0000_` reads as U+0000, and so do `_x0001_`, `_xFFFE_` and `_xFFFF_` as
 * theirs. LibreOffice drops such a character, and openpyxl leaves the escape as written. The one
 * exception is half of a surrogate pair, which reads as written unless it is the first half and
 * the escape of the second comes right after it. An escape is decoded within one run of rich
 * text: an escape, or a pair of them, that two runs split reads as written. The text of an error
 * cell is given as stored, with no escape decoded.
 *
 * Each of these types is one the writer accepts in a cell, but not every value read is written back
 * as it was: the writer refuses text longer than the 32 767 characters a cell holds in Excel, which
 * another application may have written, and drops the characters XML cannot carry.
 *
 * @example
 * for await (const { cells } of book.rows("Orders")) {
 *   const total: ReadValue = cells[2] ?? null;
 *   if (typeof total === "number") console.log(total.toFixed(2));
 * }
 *
 * @experimental
 */
export type ReadValue = string | number | boolean | null;

/**
 * One row of a sheet as `rows` yields it.
 *
 * @example
 * for await (const row of book.rows("Orders")) console.log(row.number, row.cells[0]);
 *
 * @experimental
 */
export interface ReadRow {
  /** The sheet's row number, from 1. A row without a value is not yielded, so numbers may skip. */
  readonly number: number;
  /**
   * The row's values from column A. An empty cell before a value is `null`; the array ends at the
   * row's last value, so it is never empty and never ends in `null`. A new array for every row.
   */
  readonly cells: readonly ReadValue[];
}
