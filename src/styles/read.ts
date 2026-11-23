/**
 * The read side of styles. A values-only reader needs one fact from the styles part: which cell
 * styles show a date. What decides it lives here, beside the code that writes number formats.
 */
import { XlsxReadError } from "../errors/read.js";
import type { XmlVisitor } from "../xml/read.js";
import { isBuiltInDate, showsDate } from "./read/date-format.js";

/**
 * One entry per cell style, by the index a cell's `s` holds: 1 when the style's number format shows
 * a date, 0 when it does not. An index past the end names no style. At least one entry is 1: a
 * workbook in which no cell style shows a date has no table.
 */
export type DateStyles = Uint8Array;

const NUMBER_FORMATS = 1;
const CELL_STYLES = 2;
/** The last id the standard gives a built-in format, in its Thai table, 59 to 81. */
const LAST_BUILT_IN = 81;

function damaged(): never {
  throw new XlsxReadError("INVALID_WORKBOOK", "the styles part is damaged");
}

/**
 * The visitor of the styles part. A cell style is an `xf` directly in the `cellXfs` directly under
 * the root, and its `numFmtId` names its number format, 0 when it names none; a `numFmt` directly in
 * the `numFmts` before it gives an id its code, and a code replaces a built-in id. The decision is
 * kept, not the code: each code is read once, however many styles name it, and the last `numFmt` of
 * an id wins, as LibreOffice and openpyxl read it. `applyNumberFormat` and the named styles of
 * `cellStyleXfs` change nothing, as in both. A built-in id the part does not define is the format
 * the standard gives it, or `General`.
 *
 * Refuses with `INVALID_WORKBOOK`, never repairs: a root that is not `styleSheet`; a `numFmts` or a
 * `cellXfs` stated twice, out of order or standing anywhere else, below an `AlternateContent`
 * included: LibreOffice takes the first `cellXfs` of two and openpyxl the last, and LibreOffice takes
 * a `numFmts` from the fallback of an `AlternateContent`, which openpyxl does not; anything in a
 * `numFmts` but a `numFmt` that holds nothing; anything directly in a `cellXfs` but an `xf`, and an
 * `xf` deeper in one; a `numFmt` without an id or a code; an id that is not a decimal integer; a
 * style that names an id past the built-in ones that `numFmts` does not define, for which
 * LibreOffice takes the code of a differential format in `dxfs` and openpyxl reads `General`.
 */
export function stylesVisitor(): XmlVisitor<never, DateStyles | undefined> {
  const custom = new Map<number, boolean>();
  let dates: DateStyles = new Uint8Array(64);
  let count = 0;
  let dated = false;
  let section = 0;
  let seen = 0;
  return {
    open(xml) {
      const name = xml.name;
      const depth = xml.depth;
      if (depth === 1) {
        if (name !== "styleSheet") damaged();
        return;
      }
      const listed = name === "numFmts" ? NUMBER_FORMATS : name === "cellXfs" ? CELL_STYLES : 0;
      if (listed !== 0 && (depth !== 2 || seen >= listed)) damaged();
      seen |= listed;
      if (depth === 2) {
        section = listed;
      } else if (section === NUMBER_FORMATS) {
        const id = xml.index("numFmtId");
        const code = xml.attribute("formatCode");
        if (depth !== 3 || name !== "numFmt" || !(id >= 0) || code === undefined) damaged();
        custom.set(id, showsDate(code));
      } else if (section === CELL_STYLES && (depth === 3) !== (name === "xf")) {
        damaged();
      } else if (section === CELL_STYLES && depth === 3) {
        const id = Math.max(xml.index("numFmtId"), 0);
        const date = custom.get(id);
        if (Number.isNaN(id) || (date === undefined && id > LAST_BUILT_IN)) damaged();
        if (count === dates.length) {
          const grown = new Uint8Array(count * 2);
          grown.set(dates);
          dates = grown;
        }
        const shows = date ?? isBuiltInDate(id);
        dates[count++] = shows ? 1 : 0;
        dated ||= shows;
      }
    },
    close: () => undefined,
    end: () => (dated ? dates.slice(0, count) : undefined),
  };
}
