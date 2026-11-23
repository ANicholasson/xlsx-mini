import { XlsxReadError } from "../../errors/read.js";
import type { ByteChunks } from "./contract.js";

const TAB = 9;
const LINE_FEED = 10;
const CARRIAGE_RETURN = 13;
const HASH = 35;
const AMPERSAND = 38;
const ZERO = 48;
const NINE = 57;
const LOWER_X = 120;
const BYTE_ORDER_MARK = 0xfeff;

export function refuse(what: string): never {
  throw new XlsxReadError("INVALID_XML", what);
}

export interface TextChunks {
  next(): Promise<string | undefined>;
}

export function textChunks(bytes: ByteChunks): TextChunks {
  // The decoder is told to keep the byte order mark: Node 24 and 26 drop the character after it
  // too when that is another mark and a chunk ends inside the first.
  const decoder = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true });
  let started = false;
  let ended = false;
  return {
    async next() {
      if (ended) return undefined;
      const chunk = await bytes.read();
      try {
        if (chunk === undefined) {
          ended = true;
          // The decoder refuses here a character that the last chunk left unfinished.
          return decoder.decode();
        }
        const text = decoder.decode(chunk, { stream: true });
        if (started || text === "") return text;
        started = true;
        return text.charCodeAt(0) === BYTE_ORDER_MARK ? text.slice(1) : text;
      } catch {
        return refuse("a part is not UTF-8");
      }
    },
  };
}

/**
 * A non-negative decimal integer from `text[start, end)`: `NaN` for anything but one to fifteen
 * digits. The one integer parser of the read side: row numbers, indexes and ids all use it, so
 * "a number that is not decimal" is refused the same way everywhere.
 *
 * @example parseIndex('r="41"', 3, 5) // 41
 */
export function parseIndex(text: string, start: number, end: number): number {
  if (start >= end || end - start > 15) return Number.NaN;
  let value = 0;
  for (let at = start; at < end; at++) {
    const code = text.charCodeAt(at);
    if (code < ZERO || code > NINE) return Number.NaN;
    value = value * 10 + code - ZERO;
  }
  return value;
}

/** XML 1.0 section 2.2, the production `Char`. */
function isCharacter(point: number): boolean {
  if (point < 32) return point === TAB || point === LINE_FEED || point === CARRIAGE_RETURN;
  return point < 0xd800 || (point > 0xdfff && point < 0xfffe) || point > 0xffff;
}

/**
 * The code point of the reference between `&` and `;`, or -1. XML 1.0 predefines five entities
 * (section 4.6), and a character reference must name a `Char` (4.1).
 */
export function referenced(text: string, start: number, end: number): number {
  if (text.charCodeAt(start) !== HASH) {
    const name = text.slice(start, end);
    if (name === "lt") return 60;
    if (name === "gt") return 62;
    if (name === "amp") return 38;
    if (name === "quot") return 34;
    return name === "apos" ? 39 : -1;
  }
  const radix = text.charCodeAt(start + 1) === LOWER_X ? 16 : 10;
  let at = start + (radix === 16 ? 2 : 1);
  let point = at < end ? 0 : -1;
  for (; at < end && point >= 0 && point < 0x110000; at++) {
    const code = text.charCodeAt(at);
    const letter = (code | 32) - 87;
    const digit = code >= ZERO && code <= NINE ? code - ZERO : letter > 9 ? letter : radix;
    point = digit < radix ? point * radix + digit : -1;
  }
  return point < 0x110000 && isCharacter(point) ? point : -1;
}

export const TEXT = 0;
export const CDATA = 1;
export const ATTRIBUTE = 2;
type Reading = typeof TEXT | typeof CDATA | typeof ATTRIBUTE;

/**
 * Pieces of a text joined before they are copied into one string. A string grown by `+` one piece
 * at a time is a chain of one node per piece in V8 and JavaScriptCore until it is read: 30 and 60
 * bytes for each character `decode` replaces.
 */
export const JOINED_PIECES = 1024;

/**
 * `text[start, end)` as XML 1.0 reads it: references decoded except in CDATA (4.1), line ends
 * (2.11), and in an attribute value the normalization of section 3.3.3. The result may be a slice
 * of `text`.
 */
export function decode(text: string, start: number, end: number, reading: Reading): string {
  let out = "";
  let joined = "";
  let pieces = 0;
  let copiedUpTo = start;
  for (let at = start; at < end; at++) {
    const code = text.charCodeAt(at);
    let next = at + 1;
    let read: string;
    if (code === AMPERSAND && reading !== CDATA) {
      next = text.indexOf(";", at) + 1;
      read = String.fromCodePoint(referenced(text, at + 1, next - 1));
    } else if (code === CARRIAGE_RETURN) {
      if (next < end && text.charCodeAt(next) === LINE_FEED) next++;
      read = reading === ATTRIBUTE ? " " : "\n";
    } else if (reading === ATTRIBUTE && (code === TAB || code === LINE_FEED)) {
      read = " ";
    } else {
      continue;
    }
    joined += text.slice(copiedUpTo, at) + read;
    if (++pieces === JOINED_PIECES) {
      out += detach(joined);
      joined = "";
      pieces = 0;
    }
    copiedUpTo = next;
    at = next - 1;
  }
  return out + joined + text.slice(copiedUpTo, end);
}

const PAD = "\u0001";

/**
 * A copy of `text` that does not keep the string it was cut from alive. Engines cut a string as a
 * view of its parent: V8 from 13 characters on, JavaScriptCore from fewer. A value kept from a
 * part would hold the whole decoded chunk it came from, and a column kept from a sheet the whole
 * sheet. Joined to one character the text is a new string once it is read, and JavaScriptCore
 * joins only when it is: so the character read is the number of characters to drop again.
 */
export function detach(text: string): string {
  const padded = PAD + text;
  return padded.slice(padded.charCodeAt(0));
}
