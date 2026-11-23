/**
 * From what the file stores to a `Date` whose UTC fields are the sheet's wall clock, as the writer
 * reads a `Date` (ADR 0009). The serial arithmetic is `lib/serial.ts`, which both directions share.
 */
import { epochMillisFromSerial } from "../lib/serial.js";

/**
 * The `Date` a serial names, rounded to the millisecond, or the serial itself when it names no
 * instant the writer can write: below 1 and from 60 up to 61 in the 1900 system, negative in the
 * 1904 one, past 9999-12-31 in both, each checked after the rounding.
 */
export function dateFromSerial(serial: number, date1904: boolean): Date | number {
  const millis = epochMillisFromSerial(serial, date1904);
  return Number.isNaN(millis) ? serial : new Date(millis);
}

const DASH = 45;
const POINT = 46;
const ZERO = 48;
const FIVE = 53;
const COLON = 58;
const UPPER_T = 84;
const UPPER_Z = 90;

/** The number `text[at, at + count)` writes in decimal digits, or -1. */
function digits(text: string, at: number, count: number): number {
  let value = 0;
  for (let end = at + count; at < end; at++) {
    const digit = text.charCodeAt(at) - ZERO;
    if (!(digit >= 0 && digit <= 9)) return -1;
    value = value * 10 + digit;
  }
  return value;
}

/**
 * The `Date` the text of a `t="d"` cell names, or `undefined`. Read are the forms of XML Schema's
 * `date` and `dateTime` that ECMA-376 part 1, 18.18.11, names for the type, as `2026-11-17` and
 * `2026-11-17T12:34:56`, once the white space around them is dropped: a year of four digits, seconds
 * with any number of decimals, and `Z` or no time zone, since the sheet's wall clock is UTC either
 * way. Decimals past the third round the time to the millisecond, half up, as the writer rounds a
 * serial; every decimal is read, so the work is bounded by the cell's text. Everything else is
 * `undefined`: an offset, which LibreOffice converts and openpyxl drops; a day that does not exist;
 * a year before 1900 as written, and a time that rounds up into the year 10000, which the writer
 * cannot write; `24:00:00`, a lowercase `t` or `z`, and every form XML Schema does not have.
 */
export function dateFromText(text: string): Date | undefined {
  const year = digits(text, 0, 4);
  const month = digits(text, 5, 2);
  const day = digits(text, 8, 2);
  let hour = 0;
  let minute = 0;
  let second = 0;
  let millisecond = 0;
  let up = false;
  let at = 10;
  if (text.charCodeAt(at) === UPPER_T) {
    hour = digits(text, 11, 2);
    minute = digits(text, 14, 2);
    second = digits(text, 17, 2);
    if (text.charCodeAt(13) !== COLON || text.charCodeAt(16) !== COLON) return undefined;
    at = 19;
    if (text.charCodeAt(at) === POINT) {
      const start = ++at;
      while (digits(text, at, 1) >= 0) at++;
      const read = Math.min(at - start, 3);
      if (read === 0) return undefined;
      millisecond = digits(text, start, read) * 10 ** (3 - read);
      up = at - start > 3 && text.charCodeAt(start + 3) >= FIVE;
    }
  }
  if (text.charCodeAt(at) === UPPER_Z) at++;
  if (
    at !== text.length ||
    text.charCodeAt(4) !== DASH ||
    text.charCodeAt(7) !== DASH ||
    year < 1900 ||
    (hour | minute | second) < 0 ||
    hour > 23 ||
    minute > 59 ||
    second > 59
  ) {
    return undefined;
  }
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second, millisecond));
  if (date.getUTCMonth() !== month - 1) return undefined;
  if (up) date.setTime(date.getTime() + 1);
  return date.getUTCFullYear() < 10_000 ? date : undefined;
}
