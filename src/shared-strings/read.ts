/**
 * The read side of shared strings: the table's part, and the rules of a string item, which an
 * inline string in a worksheet shares. Text is given as stored: nothing is trimmed, and only the
 * format's `_xHHHH_` escapes are decoded, as `escapeText` writes them and the applications read
 * them.
 */
import { type Meter, XlsxReadError } from "../errors/read.js";
import type { XmlCursor, XmlVisitor } from "../xml/read.js";

/**
 * The text of every entry of the table, by index; an empty entry is `""`. Every cell that names an
 * entry is given this one string.
 */
export type SharedStrings = readonly string[];

function damaged(): never {
  throw new XlsxReadError("INVALID_WORKBOOK", "the shared-string table is damaged");
}

/**
 * The visitor of the shared-string table's part. Every entry is charged to `count` before it is
 * kept, so the count the table states is never read.
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

/** What an item has held so far: its own `t`, a run, and a `t` in the run that is open. */
const ITEM_TEXT = 1;
const RUN = 2;
const RUN_TEXT = 4;

/** A string item whose own element stands at `depth`. */
export function createStringItem(depth: number): StringItem {
  let text = "";
  let inText = 0;
  let skipped = 0;
  let seen = 0;
  return {
    // Below the item's children only a run's can open: a `t` takes no child and the others are
    // passed over, so a `t` or an `rPr` two levels down stands in a run.
    open(xml) {
      if (skipped !== 0) return true;
      if (inText !== 0) return false;
      const child = xml.depth - depth === 1;
      const name = xml.name;
      if (name === "t") {
        if ((seen & (child ? ITEM_TEXT | RUN : RUN_TEXT)) !== 0) return false;
        seen |= child ? ITEM_TEXT : RUN_TEXT;
        inText = xml.depth;
      } else if (child ? name === "rPh" || name === "phoneticPr" : name === "rPr") {
        skipped = xml.depth;
      } else if (child && name === "r") {
        seen = (seen | RUN) & ~RUN_TEXT;
      } else {
        return false;
      }
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
      text += piece;
      return piece.length;
    },
    take() {
      const taken = text;
      text = "";
      seen = 0;
      return taken;
    },
  };
}

const UNDERSCORE = 95;

/**
 * Decodes the escapes `_xHHHH_`, by which the format carries a UTF-16 code unit as text: a lowercase
 * `x` and four hexadecimal digits of either case, read from left to right, so `_x0041_x0042_` is
 * `Ax0042_` and a decoded `_x005F_` starts nothing. Excel and LibreOffice decode these forms and no
 * others. Returns the text itself when it holds no `_x`.
 *
 * @example decodeEscapes("__x0078_0041_") // "_x0041_", as escapeText writes it
 */
export function decodeEscapes(text: string): string {
  let out = "";
  let copiedUpTo = 0;
  for (let at = text.indexOf("_x"); at >= 0; at = text.indexOf("_x", at)) {
    let code = 0;
    let digit = 0;
    for (let offset = 2; offset < 6 && digit >= 0; offset++) {
      digit = hexDigit(text.charCodeAt(at + offset));
      code = code * 16 + digit;
    }
    if (digit < 0 || text.charCodeAt(at + 6) !== UNDERSCORE) {
      at += 2;
      continue;
    }
    out += text.slice(copiedUpTo, at) + String.fromCharCode(code);
    copiedUpTo = at += 7;
  }
  return out + text.slice(copiedUpTo);
}

/** The value of a hexadecimal digit of either case, or -1. */
function hexDigit(code: number): number {
  if (code >= 48 && code <= 57) return code - 48;
  const letter = code | 32;
  return letter >= 97 && letter <= 102 ? letter - 87 : -1;
}
