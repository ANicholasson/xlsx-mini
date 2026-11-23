/**
 * Whether a number format shows a date. A number is a date when the format its style names shows a
 * year, a month or a day; the cell itself does not say. A time of day or an elapsed time shows none
 * of them, so such a cell stays a number, a fraction of a day.
 */

/**
 * The built-in ids that show a date when the styles part does not define them: 14 to 17 and 22 by
 * ECMA-376 part 1, 18.8.30, and the East Asian 27 to 31, 36 and 50 to 58. Ids 52 to 56 show a date
 * in some East Asian locales and a time in others, and 34 and 35 the reverse; the file does not say
 * which locale wrote it, so this reading of them is a guess (ADR 0032).
 */
export function isBuiltInDate(id: number): boolean {
  return (
    (id >= 14 && id <= 17) ||
    id === 22 ||
    (id >= 27 && id <= 31) ||
    id === 36 ||
    (id >= 50 && id <= 58)
  );
}

const QUOTE = 34;
const ASTERISK = 42;
const SEMICOLON = 59;
const OPEN_BRACKET = 91;
const BACKSLASH = 92;
const UNDERSCORE = 95;
const LOWER_A = 97;
const LOWER_D = 100;
const LOWER_G = 103;
const LOWER_H = 104;
const LOWER_M = 109;
const LOWER_S = 115;
const LOWER_Y = 121;

/** Whether `code` holds `word`, lowercase, at `at`, in any case of its letters. */
function holds(code: string, at: number, word: string): boolean {
  for (let offset = 0; offset < word.length; offset++) {
    if ((code.charCodeAt(at + offset) | 32) !== word.charCodeAt(offset)) return false;
  }
  return true;
}

/**
 * The unit of an elapsed section such as `[h]`, `[mm]` or `[ss]` between `start` and `end`: the
 * lowercase letter repeated there, or 0 for any other bracketed section.
 */
function elapsedUnit(code: string, start: number, end: number): number {
  const unit = code.charCodeAt(start) | 32;
  if (unit !== LOWER_H && unit !== LOWER_M && unit !== LOWER_S) return 0;
  for (let at = start + 1; at < end; at++) if ((code.charCodeAt(at) | 32) !== unit) return 0;
  return unit;
}

/**
 * Whether a number format code shows a year, a month or a day, read once from left to right with no
 * regular expression (ECMA-376 part 1, 18.8.31):
 *
 * - only the first section counts: a number that can be a date is positive, and a positive number
 *   shows in the first section;
 * - quoted text, the character after a backslash, an underscore or an asterisk, and bracketed
 *   sections (colours, conditions, locales) show no date; an elapsed section counts as its unit;
 * - `y` and `d` show a date in either case. `m` is a minute after an hour or before a second, and a
 *   month otherwise;
 * - `General` in any case shows a number, whatever stands beside it; the `M` of `AM/PM` is no
 *   month. Every other letter shows no date, `e`, `g` and `b` included, so the `E` of an exponent
 *   and the letters of `General` are never read as an era.
 */
export function showsDate(code: string): boolean {
  let date = false;
  let month = false;
  let afterHour = false;
  let at = 0;
  while (at < code.length) {
    const character = code.charCodeAt(at);
    if (character === SEMICOLON) break;
    if (character === QUOTE) {
      at = code.indexOf('"', at + 1) + 1;
      if (at === 0) break;
      continue;
    }
    if (character === BACKSLASH || character === UNDERSCORE || character === ASTERISK) {
      at += 2;
      continue;
    }
    let unit = character | 32;
    if (character === OPEN_BRACKET) {
      const end = code.indexOf("]", at);
      if (end < 0) break;
      unit = elapsedUnit(code, at + 1, end);
      at = end + 1;
      if (unit === LOWER_M) unit = 0;
      else if (unit === 0) continue;
    } else if (unit === LOWER_G && holds(code, at, "general")) {
      return false;
    } else if (unit === LOWER_A) {
      at += holds(code, at, "am/pm") ? 5 : 1;
      continue;
    } else if (
      unit === LOWER_Y ||
      unit === LOWER_D ||
      unit === LOWER_H ||
      unit === LOWER_M ||
      unit === LOWER_S
    ) {
      while ((code.charCodeAt(at) | 32) === unit) at++;
    } else {
      at++;
      continue;
    }
    if (month && unit !== LOWER_S) date = true;
    month = unit === LOWER_M && !afterHour;
    afterHour = unit === LOWER_H;
    if (unit === LOWER_Y || unit === LOWER_D) date = true;
  }
  return date || month;
}
