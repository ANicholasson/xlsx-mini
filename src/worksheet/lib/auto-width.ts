import type { CellStyle } from "../../styles/index.js";
import { MAX_COLUMNS } from "./cell-ref.js";

/**
 * Excel measures a column in digits of the default font, Calibri 11 at 96 dpi: a digit is 7 pixels
 * and a column carries 5 pixels of padding. A fitted column adds 7 more around its widest cell.
 */
const DIGIT_PIXELS = 7;
const FIT_PADDING = 12;
const MIN_WIDTH = 3;
const MAX_WIDTH = 60;
/** Past this many pixels a column is at the ceiling and measures nothing more. */
const CEILING_PIXELS = MAX_WIDTH * DIGIT_PIXELS - FIT_PADDING;
const DEFAULT_FONT_SIZE = 11;
const BOLD_FACTOR = 1.1;
/** A character the table does not cover, a wide (East Asian or emoji) character, and the widest tabled one. */
const OTHER_PIXELS = 8;
const WIDE_PIXELS = 14;
const WIDEST_PIXELS = 14;
/** The General format shows at most this many characters before it shortens a number. */
const GENERAL_MAX_CHARS = 11;
const GENERAL = 0;

/** Pixel widths of the printable ASCII characters, space to tilde, in Calibri 11 at 96 dpi. */
// prettier-ignore
const ASCII_PIXELS = new Uint8Array([
  3, 5, 6, 7, 7, 11, 10, 3, 5, 5, 7, 7, 4, 5, 4, 6, // space ! " # $ % & ' ( ) * + , - . /
  7, 7, 7, 7, 7, 7, 7, 7, 7, 7, // 0-9
  4, 4, 7, 7, 7, 7, 13, // : ; < = > ? @
  9, 8, 8, 9, 7, 7, 9, 9, 4, 5, 8, 6, 12, 10, 10, 8, 10, 8, 7, 7, 9, 9, 13, 8, 7, 7, // A-Z
  5, 6, 5, 7, 7, 4, // [ \ ] ^ _ `
  7, 8, 6, 8, 8, 5, 7, 8, 3, 4, 7, 3, 12, 8, 8, 8, 8, 5, 6, 5, 8, 7, 11, 7, 7, 6, // a-z
  5, 7, 5, 7, // { | } ~
]);

/** The pixels one UTF-16 code unit takes; a surrogate pair counts once, a combining mark not at all. */
function pixelsOf(code: number): number {
  if (code < 32) return 0;
  if (code < 127) return ASCII_PIXELS[code - 32]!;
  if (code >= 0x0300 && code <= 0x036f) return 0;
  if (code >= 0xdc00 && code <= 0xdfff) return 0;
  if (
    (code >= 0xd800 && code <= 0xdbff) ||
    (code >= 0x1100 && code <= 0x115f) ||
    (code >= 0x2e80 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7a3) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe4f) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  ) {
    return WIDE_PIXELS;
  }
  return OTHER_PIXELS;
}

/** The pixels of the widest line of a text. */
function textPixels(text: string): number {
  let widest = 0;
  let line = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 10) {
      if (line > widest) widest = line;
      line = 0;
    } else {
      line += pixelsOf(code);
    }
  }
  return line > widest ? line : widest;
}

/** What a number format says about the length of what it shows, read once per format. */
interface FormatShape {
  /** The pixels of the date and time tokens, or 0 for a numeric format. */
  readonly datePixels: number;
  /** The fewest integer digits the format shows: its `0` placeholders before the point. */
  readonly integerMin: number;
  readonly grouping: boolean;
  readonly decimals: number;
  /** What the value is multiplied by before it is shown: 100 per `%`, a thousandth per trailing comma. */
  readonly scale: number;
  /** The pixels of an `E+00` exponent, or 0; with an exponent the mantissa has `integerMin` digits. */
  readonly exponentPixels: number;
  /** The pixels of the literal text around the digits. */
  readonly literalPixels: number;
}

const GENERAL_SHAPE: FormatShape = {
  datePixels: 0,
  integerMin: 1,
  grouping: false,
  decimals: 0,
  scale: 1,
  exponentPixels: 0,
  literalPixels: 0,
};

/** The built-in number formats that show something other than General, by id. */
const BUILT_IN_FORMATS: Readonly<Record<number, string>> = {
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
};

/** The characters a run of one date letter shows: `mmmm` is a month name, `yyyy` a year, the rest two digits. */
function dateTokenChars(letter: string, run: number): number {
  if (letter === "m" || letter === "d") {
    if (run >= 5) return 1;
    if (run === 4) return 9;
    if (run === 3) return 3;
    return 2;
  }
  if (letter === "y" || letter === "e" || letter === "b") return run >= 3 ? 4 : 2;
  return 2;
}

const DATE_LETTERS = new Set(["y", "m", "d", "h", "s", "e", "b", "g"]);
const PLACEHOLDERS = new Set(["0", "#", "?"]);

/**
 * Reads the first section of a format string for the width of what it shows. The format is read,
 * not applied: placeholders say how many digits come out, tokens how many letters a date takes,
 * and everything else is literal text. Unknown constructs count as literal, which errs wide.
 */
function readFormat(format: string): FormatShape {
  const section = format.indexOf(";");
  if ((section === -1 ? format : format.slice(0, section)).trim().toLowerCase() === "general") {
    return GENERAL_SHAPE;
  }
  let datePixels = 0;
  let integerMin = 0;
  let grouping = false;
  let decimals = 0;
  let scale = 1;
  let exponentPixels = 0;
  let literalPixels = 0;
  let afterPoint = false;
  let afterExponent = false;
  let exponentDigits = 0;
  let i = 0;
  while (i < format.length) {
    const char = format[i]!;
    const lower = char.toLowerCase();
    if (char === ";") break;
    if (char === '"') {
      const end = format.indexOf('"', i + 1);
      const stop = end === -1 ? format.length : end;
      literalPixels += textPixels(format.slice(i + 1, stop));
      i = stop + 1;
    } else if (char === "\\") {
      if (i + 1 < format.length) literalPixels += pixelsOf(format.charCodeAt(i + 1));
      i += 2;
    } else if (char === "_") {
      literalPixels += ASCII_PIXELS[0]!;
      i += 2;
    } else if (char === "*") {
      i += 2;
    } else if (char === "[") {
      const end = format.indexOf("]", i + 1);
      const stop = end === -1 ? format.length : end;
      const inside = format.slice(i + 1, stop).toLowerCase();
      // Elapsed time such as [h] shows as many digits as it needs; one more than the clock form.
      if (/^(h+|m+|s+)$/.test(inside)) datePixels += (inside.length + 1) * DIGIT_PIXELS;
      i = stop + 1;
    } else if (PLACEHOLDERS.has(char)) {
      if (afterExponent) exponentDigits++;
      else if (afterPoint) decimals++;
      else if (char === "0") integerMin++;
      i++;
    } else if (char === ".") {
      if (!afterPoint && !afterExponent) afterPoint = true;
      else literalPixels += pixelsOf(46);
      i++;
    } else if (char === ",") {
      if (!afterPoint && !afterExponent) {
        if (i + 1 < format.length && PLACEHOLDERS.has(format[i + 1]!)) grouping = true;
        else scale /= 1000;
      } else {
        literalPixels += pixelsOf(44);
      }
      i++;
    } else if (char === "%") {
      scale *= 100;
      literalPixels += pixelsOf(37);
      i++;
    } else if (lower === "e" && (format[i + 1] === "+" || format[i + 1] === "-")) {
      afterExponent = true;
      afterPoint = false;
      i += 2;
    } else if (lower === "a" && format.slice(i, i + 5).toUpperCase() === "AM/PM") {
      datePixels += 2 * DIGIT_PIXELS;
      i += 5;
    } else if (lower === "a" && format.slice(i, i + 3).toUpperCase() === "A/P") {
      datePixels += DIGIT_PIXELS;
      i += 3;
    } else if (DATE_LETTERS.has(lower)) {
      let run = 1;
      while (i + run < format.length && format[i + run]!.toLowerCase() === lower) run++;
      datePixels += dateTokenChars(lower, run) * DIGIT_PIXELS;
      i += run;
    } else if (char === "@") {
      i++;
    } else {
      literalPixels += pixelsOf(format.charCodeAt(i));
      i++;
    }
  }
  if (afterExponent) exponentPixels = (2 + exponentDigits) * DIGIT_PIXELS;
  return {
    datePixels,
    integerMin: integerMin === 0 ? 1 : integerMin,
    grouping,
    decimals,
    scale,
    exponentPixels,
    literalPixels,
  };
}

/** The characters the General format shows for a number: its text, shortened past eleven. */
function generalChars(value: number): number {
  if (Number.isInteger(value) && Math.abs(value) < 1e11) {
    const magnitude = Math.abs(value);
    const digits = magnitude < 1 ? 1 : Math.floor(Math.log10(magnitude) + 1e-10) + 1;
    return value < 0 ? digits + 1 : digits;
  }
  const length = String(value).length;
  return length > GENERAL_MAX_CHARS ? GENERAL_MAX_CHARS : length;
}

/** The pixels a number takes under a format, from the format's shape and the value's magnitude. */
function numberPixels(value: number, shape: FormatShape): number {
  if (shape.datePixels > 0) return shape.datePixels + shape.literalPixels;
  if (shape === GENERAL_SHAPE) return generalChars(value) * DIGIT_PIXELS;
  let integerDigits = shape.integerMin;
  if (shape.exponentPixels === 0) {
    const magnitude = Math.abs(value) * shape.scale;
    const digits = magnitude < 1 ? 1 : Math.floor(Math.log10(magnitude) + 1e-10) + 1;
    if (digits > integerDigits) integerDigits = digits;
  }
  let chars = integerDigits;
  if (shape.grouping) chars += Math.floor((integerDigits - 1) / 3);
  if (shape.decimals > 0) chars += shape.decimals + 1;
  if (value < 0) chars++;
  return chars * DIGIT_PIXELS + shape.exponentPixels + shape.literalPixels;
}

/**
 * Measures the cells of the columns that size themselves, one pixel count per column, and turns
 * them into widths at the end. Nothing is allocated per cell: a text that cannot beat the column's
 * current maximum is not even scanned, and a column at the ceiling is left alone.
 */
export interface AutoWidths {
  /** Whether the cell at this position is measured: an auto-sized column, below the ceiling, not the start of a merge across columns. */
  wants(column: number, rowIndex: number): boolean;
  /** Records a text cell, a boolean as its `TRUE` or `FALSE`, or a cached formula text. */
  text(
    column: number,
    text: string,
    cell: CellStyle | undefined,
    columnStyle: CellStyle | undefined,
  ): void;
  /** Records a number or a date serial under the number format the cell resolves to, or `dateDefault`. */
  number(
    column: number,
    value: number,
    dateDefault: number,
    cell: CellStyle | undefined,
    columnStyle: CellStyle | undefined,
  ): void;
  /** Ends measuring: the widths stay as they are, whatever rows follow. */
  stop(): void;
  /** The width per column, in characters, or undefined where nothing was measured. */
  widths(): readonly (number | undefined)[];
}

/** One measured column: its widest cell so far and what its own style settles once. */
interface Slot {
  pixels: number;
  /** The factor for cells that bring no style of their own, or -1 until the column style is seen. */
  factor: number;
  /** The shape per date default the column's format resolves to, filled on first use. */
  readonly shapes: (FormatShape | undefined)[];
  /**
   * Per date default, the largest magnitude measured under that shape: a smaller one cannot widen
   * the column, since a format's width grows with the magnitude. General is the exception, where
   * only whole numbers grow that way; `0.123456789` is wider than `1000`.
   */
  readonly largest: number[];
}

/**
 * The measurer for a sheet: `measured[i]` says whether column `i` sizes itself, `beyond` whether
 * columns past that list do, and `mergeStarts` holds `rowIndex * MAX_COLUMNS + column` for every
 * cell that starts a merge across more than one column.
 */
export function createAutoWidths(
  measured: readonly boolean[],
  beyond: boolean,
  mergeStarts: ReadonlySet<number> | undefined,
): AutoWidths {
  /** A slot per column that sizes itself, null for one that does not, undefined until first seen. */
  const slots: (Slot | null | undefined)[] = [];
  const shapes = new Map<string | number, FormatShape>();
  let active = true;

  const shapeOf = (format: string | number | undefined, dateDefault: number): FormatShape => {
    const key = format ?? dateDefault;
    if (key === GENERAL) return GENERAL_SHAPE;
    let shape = shapes.get(key);
    if (shape === undefined) {
      const text = typeof key === "number" ? BUILT_IN_FORMATS[key] : key;
      shape = text === undefined ? GENERAL_SHAPE : readFormat(text);
      shapes.set(key, shape);
    }
    return shape;
  };

  const factorOf = (cell: CellStyle | undefined, columnStyle: CellStyle | undefined): number => {
    const size = cell?.fontSize ?? columnStyle?.fontSize;
    const bold = cell?.bold ?? columnStyle?.bold;
    let factor = size === undefined ? 1 : size / DEFAULT_FONT_SIZE;
    if (bold === true) factor *= BOLD_FACTOR;
    return factor;
  };

  const slotFactor = (slot: Slot, columnStyle: CellStyle | undefined): number => {
    if (slot.factor < 0) slot.factor = factorOf(undefined, columnStyle);
    return slot.factor;
  };

  return {
    wants(column, rowIndex) {
      if (!active) return false;
      let slot = slots[column];
      if (slot === undefined) {
        const wanted = column < measured.length ? measured[column] === true : beyond;
        slot = wanted ? { pixels: 0, factor: -1, shapes: [], largest: [] } : null;
        slots[column] = slot;
      }
      if (slot === null || slot.pixels >= CEILING_PIXELS) return false;
      return mergeStarts === undefined || !mergeStarts.has(rowIndex * MAX_COLUMNS + column);
    },
    text(column, text, cell, columnStyle) {
      const slot = slots[column]!;
      const factor =
        cell === undefined ? slotFactor(slot, columnStyle) : factorOf(cell, columnStyle);
      if (text.length * WIDEST_PIXELS * factor <= slot.pixels) return;
      const measuredPixels = textPixels(text) * factor;
      if (measuredPixels > slot.pixels) slot.pixels = measuredPixels;
    },
    number(column, value, dateDefault, cell, columnStyle) {
      const slot = slots[column]!;
      let measuredPixels: number;
      if (cell !== undefined) {
        const shape = shapeOf(cell.numberFormat ?? columnStyle?.numberFormat, dateDefault);
        measuredPixels = numberPixels(value, shape) * factorOf(cell, columnStyle);
      } else {
        let shape = slot.shapes[dateDefault];
        if (shape === undefined) {
          shape = shapeOf(columnStyle?.numberFormat, dateDefault);
          slot.shapes[dateDefault] = shape;
        }
        // Under one shape the width grows with the magnitude, a negative counting one digit more,
        // so a value no larger than the largest measured cannot widen the column. Under General
        // that holds for whole numbers only; a fraction is always measured.
        if (shape !== GENERAL_SHAPE || Number.isInteger(value)) {
          const magnitude = value < 0 ? -value * 10 : value;
          if (magnitude <= (slot.largest[dateDefault] ?? -1)) return;
          slot.largest[dateDefault] = magnitude;
        }
        measuredPixels = numberPixels(value, shape) * slotFactor(slot, columnStyle);
      }
      if (measuredPixels > slot.pixels) slot.pixels = measuredPixels;
    },
    stop() {
      active = false;
    },
    widths() {
      const widths: (number | undefined)[] = [];
      for (let column = 0; column < slots.length; column++) {
        const slot = slots[column];
        if (slot === undefined || slot === null || slot.pixels <= 0) {
          widths.push(undefined);
          continue;
        }
        const width = Math.round(((slot.pixels + FIT_PADDING) / DIGIT_PIXELS) * 100) / 100;
        widths.push(Math.min(MAX_WIDTH, Math.max(MIN_WIDTH, width)));
      }
      return widths;
    },
  };
}
