/**
 * The read side of shared strings: the table, and the rules of a string item, which an inline
 * string in a worksheet shares. Text is given as stored: nothing is trimmed.
 */
import { type Meter, XlsxReadError } from "../errors/read.js";
import { JOINED_PIECES, type XmlCursor, type XmlVisitor, detach } from "../xml/read.js";

/**
 * The text of every entry of the table, by index; an empty entry is `""`. A cell that names an
 * entry is given the entry's string itself, never a copy.
 */
export type SharedStrings = readonly string[];

function damaged(): never {
  throw new XlsxReadError("INVALID_WORKBOOK", "the shared-string table is damaged");
}

/**
 * Every entry is charged to `count` before it is kept, so the count the table states is never read.
 *
 * Refuses with `INVALID_WORKBOOK` a part whose root is not `sst`, and anything under the root but
 * an `si` and the extensions of `extLst`, which may hold no `si`: an entry in a markup-compatibility
 * block or an extension is counted by one reader and passed over by another, and every index after
 * it would name another string.
 */
export function sharedStringsVisitor(
  count: Meter<"sharedStrings">,
): XmlVisitor<never, SharedStrings> {
  const strings: string[] = [];
  const item = createStringItem(2);
  let skipped = 0;
  return {
    open(xml) {
      const depth = xml.depth;
      const name = xml.name;
      if (skipped !== 0) {
        if (name === "si") damaged();
      } else if (depth > 2 ? item.open(xml) : depth === 1 ? name === "sst" : name === "si") {
        if (depth === 2) count.charge(1);
      } else if (depth === 2 && name === "extLst") {
        skipped = 2;
      } else {
        damaged();
      }
    },
    close(xml) {
      const depth = xml.depth;
      if (skipped !== 0) {
        if (depth === skipped) skipped = 0;
      } else if (depth > 2) {
        item.close(xml);
      } else if (depth === 2) {
        strings.push(item.take());
      }
      return undefined;
    },
    end: () => strings,
  };
}

/**
 * The text of one string item, `si` in the table or `is` in a cell, told of every element below the
 * item's own. The item's text is the text of its `t` and of the `t` of each rich run `r`, in order,
 * each with its escapes decoded; the properties of a run (`rPr`) and the phonetic runs (`rPh`,
 * `phoneticPr`) hold no text of it. One state object serves every item of a part.
 */
interface StringItem {
  /**
   * An element opens below the item's own. False when an item cannot hold it there: anything but
   * the elements above; any element inside a `t`, whose text would then lose characters; and a
   * second `t` in the item or in a run, or a `t` of the item after a run, which LibreOffice and
   * openpyxl read in different ways.
   */
  open(xml: XmlCursor): boolean;
  /** An element closes below the item's own. Returns the number of characters it added. */
  close(xml: XmlCursor): number;
  /** The item's text, and starts the next item. */
  take(): string;
}

const ITEM_TEXT = 1;
const RUN = 2;
const RUN_TEXT = 4;

/** A string item whose own element stands at `depth`. */
export function createStringItem(depth: number): StringItem {
  let text = "";
  let joined = "";
  let pieces = 0;
  let inText = 0;
  let skipped = 0;
  let seen = 0;
  return {
    open(xml) {
      if (skipped !== 0) return true;
      if (inText !== 0) return false;
      const name = xml.name;
      const ownChild = xml.depth - depth === 1;
      if (ownChild) {
        if (name === "r") {
          seen = (seen | RUN) & ~RUN_TEXT;
          return true;
        }
        if (name === "rPh" || name === "phoneticPr") {
          skipped = xml.depth;
          return true;
        }
        if (name !== "t" || (seen & (ITEM_TEXT | RUN)) !== 0) return false;
        seen |= ITEM_TEXT;
      } else {
        if (name === "rPr") {
          skipped = xml.depth;
          return true;
        }
        if (name !== "t" || (seen & RUN_TEXT) !== 0) return false;
        seen |= RUN_TEXT;
      }
      inText = xml.depth;
      return true;
    },
    close(xml) {
      const at = xml.depth;
      if (skipped !== 0) {
        if (at === skipped) skipped = 0;
        return 0;
      }
      if (at !== inText) return 0;
      inText = 0;
      const piece = decodeEscapes(xml.text());
      joined += piece;
      if (++pieces === JOINED_PIECES) {
        text += detach(joined);
        joined = "";
        pieces = 0;
      }
      return piece.length;
    },
    take() {
      const taken = text === "" && pieces < 2 ? joined : detach(text + joined);
      text = "";
      joined = "";
      pieces = 0;
      seen = 0;
      return taken;
    },
  };
}

const UNDERSCORE = 95;
const LOWER_X = 120;

/**
 * Decodes the escapes `_xHHHH_`, by which the format carries a UTF-16 code unit as text: a lowercase
 * `x` and four hexadecimal digits of either case, read from left to right, so `_x0041_x0042_` is
 * `Ax0042_` and a decoded `_x005F_` starts nothing. Excel and LibreOffice decode these forms and no
 * others. Half of a surrogate pair is decoded only with the escape of its other half after it, and
 * kept as text otherwise, as openpyxl keeps every escape: LibreOffice ends the text there. Returns
 * the text itself when it decodes no escape, and a copy of its own otherwise.
 *
 * @example decodeEscapes("__x0078_0041_") // "_x0041_", as escapeText writes it
 */
export function decodeEscapes(text: string): string {
  let out = "";
  let copiedUpTo = 0;
  for (let at = text.indexOf("_x"); at >= 0; at = text.indexOf("_x", at)) {
    const code = escapeAt(text, at);
    const low = code >= 0xd800 && code < 0xdc00 ? escapeAt(text, at + 7) : -1;
    const paired = low >= 0xdc00 && low < 0xe000;
    if (code < 0 || (code >= 0xd800 && code < 0xe000 && !paired)) {
      at += 2;
      continue;
    }
    out +=
      text.slice(copiedUpTo, at) +
      String.fromCharCode(code) +
      (paired ? String.fromCharCode(low) : "");
    copiedUpTo = at += paired ? 14 : 7;
  }
  return copiedUpTo === 0 ? text : detach(out + text.slice(copiedUpTo));
}

function escapeAt(text: string, at: number): number {
  if (text.charCodeAt(at + 6) !== UNDERSCORE || text.charCodeAt(at + 1) !== LOWER_X) return -1;
  let code = 0;
  for (let offset = 2; offset < 6; offset++) {
    const digit = hexDigit(text.charCodeAt(at + offset));
    if (digit < 0) return -1;
    code = code * 16 + digit;
  }
  return text.charCodeAt(at) === UNDERSCORE ? code : -1;
}

function hexDigit(code: number): number {
  if (code >= 48 && code <= 57) return code - 48;
  const letter = code | 32;
  return letter >= 97 && letter <= 102 ? letter - 87 : -1;
}
