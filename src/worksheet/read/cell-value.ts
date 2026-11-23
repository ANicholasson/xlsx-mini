import { type SharedStrings, decodeEscapes } from "../../shared-strings/read.js";
import { parseIndex } from "../../xml/read.js";
import type { ReadValue } from "./model.js";

/** The type of a cell, from its `t` attribute, as a small number compared once per cell. */
const NUMBER = 0;
const SHARED_STRING = 1;
const BOOLEAN = 2;
const ERROR = 3;
export const DATE = 4;
const FORMULA_STRING = 5;
export const INLINE_STRING = 6;
export const UNKNOWN = -1;

const ONE_LETTER = "nsbed";

/**
 * The type the `t` value `buffer[start, end)` names, compared in place: `n`, `s`, `b`, `e`, `d`,
 * `str` or `inlineStr`, else `UNKNOWN`. A cell without the attribute (`start` is -1) is a number.
 */
export function cellKindOf(buffer: string, start: number, end: number): number {
  if (start < 0) return NUMBER;
  const length = end - start;
  if (length === 1) return ONE_LETTER.indexOf(buffer.charAt(start));
  if (length === 3 && buffer.startsWith("str", start)) return FORMULA_STRING;
  return length === 9 && buffer.startsWith("inlineStr", start) ? INLINE_STRING : UNKNOWN;
}

/**
 * The value of a cell from its type and what it stores: the text of its `v`, or of its inline
 * string, or `undefined` when it stores neither, which is `null`. Empty text is `null`. `undefined`
 * when the text is not a value of that type: a number that is not decimal or not finite, a boolean
 * other than `0` or `1`, an index that names no shared string, each of them empty too, and every
 * date stored as text.
 */
export function cellValue(
  kind: number,
  stored: string | undefined,
  strings: SharedStrings,
): ReadValue | undefined {
  if (stored === undefined) return null;
  switch (kind) {
    case NUMBER:
      return numberOf(stored);
    case SHARED_STRING: {
      const text = strings[parseIndex(stored, 0, stored.length)];
      return text === undefined ? undefined : text || null;
    }
    case BOOLEAN:
      return stored === "1" ? true : stored === "0" ? false : undefined;
    case FORMULA_STRING:
      return decodeEscapes(stored) || null;
    case ERROR:
    case INLINE_STRING:
      return stored || null;
    default:
      return undefined;
  }
}

const TAB = 9;
const LINE_FEED = 10;
const CARRIAGE_RETURN = 13;
const SPACE = 32;
const PLUS = 43;
const MINUS = 45;
const POINT = 46;
const ZERO = 48;
const NINE = 57;
const LOWER_E = 101;

function isSpace(code: number): boolean {
  return code === SPACE || code === TAB || code === LINE_FEED || code === CARRIAGE_RETURN;
}

function isDigit(code: number): boolean {
  return code >= ZERO && code <= NINE;
}

/**
 * The number a `v` holds, as XML Schema's `double` spells one: a sign, digits with a decimal point
 * anywhere in them, an exponent, and white space at either end. The spellings of an infinity and
 * of NaN, and a value too large to be finite, are no number; so are a hexadecimal or binary prefix
 * and a comma, which `Number` or LibreOffice would read as something.
 */
function numberOf(text: string): number | undefined {
  let at = 0;
  let end = text.length;
  while (isSpace(text.charCodeAt(at))) at++;
  while (end > at && isSpace(text.charCodeAt(end - 1))) end--;
  let code = text.charCodeAt(at);
  if (code === PLUS || code === MINUS) code = text.charCodeAt(++at);
  let digits = 0;
  for (; isDigit(code); digits++) code = text.charCodeAt(++at);
  if (code === POINT) {
    for (code = text.charCodeAt(++at); isDigit(code); digits++) code = text.charCodeAt(++at);
  }
  if (digits === 0) return undefined;
  if ((code | SPACE) === LOWER_E) {
    code = text.charCodeAt(++at);
    if (code === PLUS || code === MINUS) code = text.charCodeAt(++at);
    while (isDigit(code)) code = text.charCodeAt(++at);
  }
  const value = Number(text);
  return at === end && Number.isFinite(value) ? value : undefined;
}
