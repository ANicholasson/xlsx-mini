export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;

const letters: string[] = [];

/**
 * Excel column letters for a zero-based column index: 0 is `A`, 25 is `Z`, 26 is `AA`.
 * Results are cached, so a wide sheet pays the conversion once per column.
 */
export function columnLetter(index: number): string {
  const cached = letters[index];
  if (cached !== undefined) return cached;
  let result = "";
  let remaining = index;
  do {
    result = String.fromCharCode(65 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  letters[index] = result;
  return result;
}

/** Zero-based column index of column letters such as `A`, `Z` or `AA`. */
function columnIndex(letters: string): number {
  let index = 0;
  for (let i = 0; i < letters.length; i++) index = index * 26 + (letters.charCodeAt(i) - 64);
  return index - 1;
}

/** A parsed reference: zero-based columns, one-based rows, inclusive on both ends. */
export interface Box {
  /** The reference as it is written to the file: `B2`, `A1:D4`, or `B1:B1048576` for a whole column. */
  readonly ref: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

const CELL = /^([A-Z]{1,3})([1-9]\d{0,6})$/;
const RANGE = /^([A-Z]{1,3})([1-9]\d{0,6}):([A-Z]{1,3})([1-9]\d{0,6})$/;
const COLUMNS = /^([A-Z]{1,3}):([A-Z]{1,3})$/;

/**
 * Parses a cell such as `B2`, a range such as `A1:D4` or whole columns such as `B:B` into a box
 * or returns undefined when the text is none of those; `boxProblem` then says whether it fits the
 * sheet. Whole columns become the rows 1 to 1 048 576, which is how Excel stores them. Absolute
 * markers (`$`) are not accepted: the writer adds them where a part needs them.
 */
export function parseReference(ref: unknown): Box | undefined {
  if (typeof ref !== "string") return undefined;
  let left: number;
  let right: number;
  let top: number;
  let bottom: number;
  let written = ref;
  const range = RANGE.exec(ref);
  if (range !== null) {
    left = columnIndex(range[1]!);
    top = Number(range[2]);
    right = columnIndex(range[3]!);
    bottom = Number(range[4]);
  } else {
    const cell = CELL.exec(ref);
    if (cell !== null) {
      left = right = columnIndex(cell[1]!);
      top = bottom = Number(cell[2]);
    } else {
      const columns = COLUMNS.exec(ref);
      if (columns === null) return undefined;
      left = columnIndex(columns[1]!);
      right = columnIndex(columns[2]!);
      top = 1;
      bottom = MAX_ROWS;
      written = `${columns[1]}1:${columns[2]}${MAX_ROWS}`;
    }
  }
  return { ref: written, top, left, bottom, right };
}

/** Why a parsed box cannot be used as it is, or undefined when it lies inside the sheet the right way round. */
export function boxProblem(box: Box): string | undefined {
  if (box.right >= MAX_COLUMNS || box.bottom > MAX_ROWS)
    return `Range ${box.ref} lies outside the sheet`;
  if (box.right < box.left || box.bottom < box.top) {
    return `Range ${box.ref} must run from its top-left cell to its bottom-right cell`;
  }
  return undefined;
}

const DIGIT_0 = 48;
const DIGIT_9 = 57;

/**
 * `A1:C4` as `$A$1:$C$4`, the absolute form defined names use: a `$` goes before every run of
 * letters and every run of digits. A scan rather than a pattern, since the text may come from the
 * caller and a backtracking pattern can be made slow.
 */
export function absolute(ref: string): string {
  let out = "";
  let wasDigit = false;
  let wasLetter = false;
  for (let i = 0; i < ref.length; i++) {
    const code = ref.charCodeAt(i);
    const digit = code >= DIGIT_0 && code <= DIGIT_9;
    const letter = !digit && code !== 58;
    if ((letter && !wasLetter) || (digit && !wasDigit)) out += "$";
    out += ref[i];
    wasDigit = digit;
    wasLetter = letter;
  }
  return out;
}

/** A sheet name as a formula reference: quoted, with apostrophes doubled. */
export function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}
