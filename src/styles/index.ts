import { XlsxMiniError, describeValue } from "../errors/index.js";
import { XML_DECLARATION, escapeXml } from "../xml/index.js";

/** Line style of a cell border. */
export type BorderStyle = "thin" | "medium" | "thick";

/** A border per side; a side left out has no border. */
export interface Borders {
  readonly top?: BorderStyle;
  readonly right?: BorderStyle;
  readonly bottom?: BorderStyle;
  readonly left?: BorderStyle;
}

export type HorizontalAlign = "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";

/**
 * How a cell looks. Set on a cell object or as a column's default; a cell's own properties win,
 * property by property. Colours are `#rrggbb`.
 *
 * @example
 * const money: CellStyle = { numberFormat: '#,##0.00 "kr"', align: "right" };
 * const header: CellStyle = { bold: true, fill: "#d9e1f2", border: "thin" };
 */
export interface CellStyle {
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** Font size in points, from 1 to 409; the default is 11. */
  readonly fontSize?: number;
  /** Font colour as `#rrggbb`. */
  readonly fontColor?: string;
  /** Solid background colour as `#rrggbb`. */
  readonly fill?: string;
  /** One line style for all four sides, or a style per side. */
  readonly border?: BorderStyle | Borders;
  readonly align?: HorizontalAlign;
  readonly verticalAlign?: VerticalAlign;
  /** Wrap long text onto several lines inside the cell. */
  readonly wrap?: boolean;
  /**
   * Indent in Excel's indent steps, from 0 to 250. Excel indents left- or right-aligned text, so an
   * indent without `align` aligns left, and an indent with `center` is an error.
   */
  readonly indent?: number;
  /** An Excel format string such as `"0.0%"` or `"yyyy-mm-dd"`, or a built-in format id from 0 to 49. */
  readonly numberFormat?: string | number;
}

/** Hands out cell-format indexes, deduplicated, and renders `xl/styles.xml` with exactly what was used. */
export interface StyleTable {
  /**
   * The index of the cell format for a cell style merged over a column style. `dateDefault` is the
   * built-in number format a date-like value falls back to when neither style sets one, or 0.
   * Throws with the offending property as the path, for the caller to prefix.
   */
  indexFor(cell: CellStyle | undefined, column: CellStyle | undefined, dateDefault: number): number;
  /** Throws for a value Excel would refuse, with the property name as the path, registering nothing. */
  validate(style: CellStyle): void;
  /** Renders the styles part. Call once, after every sheet has been rendered. */
  render(): string;
}

const GENERAL = 0;
const MAX_BUILT_IN_FORMAT = 49;
const FIRST_CUSTOM_FORMAT = 164;
const MAX_FORMAT_LENGTH = 255;
const DEFAULT_FONT_SIZE = 11;
const MIN_FONT_SIZE = 1;
const MAX_FONT_SIZE = 409;
const MAX_INDENT = 250;
const COLOR = /^#[0-9a-fA-F]{6}$/;
const BORDER_STYLES: ReadonlySet<string> = new Set(["thin", "medium", "thick"]);
const HORIZONTAL: ReadonlySet<string> = new Set(["left", "center", "right"]);
const VERTICAL: ReadonlySet<string> = new Set(["top", "middle", "bottom"]);
const SIDES = ["left", "right", "top", "bottom"] as const;
type Sides = Readonly<Record<(typeof SIDES)[number], BorderStyle | undefined>>;

const DEFAULT_FONT = '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>';
const DEFAULT_FILLS =
  '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>';
const DEFAULT_BORDER = "<border><left/><right/><top/><bottom/><diagonal/></border>";

/** The parts of a cell format after merging and validation; every field at its default is 0 or "". */
interface Resolved {
  readonly numberFormat: number;
  readonly fontId: number;
  readonly fillId: number;
  readonly borderId: number;
  readonly align: string;
  readonly verticalAlign: string;
  readonly wrap: boolean;
  readonly indent: number;
}

function invalid(property: string, message: string): XlsxMiniError {
  return new XlsxMiniError("INVALID_STYLE", property, message);
}

/** The style itself must be an object; the path is empty because the whole value is wrong. */
function styleObject(style: unknown): void {
  if (typeof style !== "object" || style === null) {
    throw invalid("", `A style is an object, not ${describeValue(style)}`);
  }
}

function flag(property: string, value: unknown): boolean {
  if (value === undefined) return false;
  if (typeof value !== "boolean") {
    throw invalid(property, `${property} is true or false, not ${describeValue(value)}`);
  }
  return value;
}

/** `#rrggbb` as the ARGB hex Excel stores, or an error naming the property. */
function argb(property: string, color: string): string {
  if (typeof color !== "string" || !COLOR.test(color)) {
    throw invalid(property, `Colours are written as #rrggbb, not ${describeValue(color)}`);
  }
  return `FF${color.slice(1).toUpperCase()}`;
}

function borderSide(property: string, style: unknown): BorderStyle | undefined {
  if (style === undefined) return undefined;
  if (typeof style !== "string" || !BORDER_STYLES.has(style)) {
    throw invalid(property, `Border styles are thin, medium or thick, not ${describeValue(style)}`);
  }
  return style as BorderStyle;
}

/**
 * Index 0 is always the default format. Fonts, fills, borders, custom number formats and cell
 * formats are each registered once per distinct value. A style object seen again, alone or over the
 * same column style, is resolved through WeakMaps by identity, so a column or header style reused
 * across thousands of cells costs lookups and no string per cell. Excel expects fill 0 to be
 * `none`, fill 1 to be `gray125`, and a `Normal` cell style.
 */
export function createStyleTable(): StyleTable {
  const cellFormats: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const byKey = new Map<string, number>();
  const bySingleObject = new WeakMap<CellStyle, number>();
  /** The same, per date default, for a lone style on date cells: no key string per cell. */
  const bySingleDated = new WeakMap<CellStyle, number[]>();
  /** Cell style over column style, per date default, by the identity of both objects. */
  const byPair = new WeakMap<CellStyle, WeakMap<CellStyle, number[]>>();
  /** Cell-format index per built-in date format, for unstyled date cells: no key string per cell. */
  const byDateDefault: number[] = [];
  const customFormats = new Map<string, number>();
  const fonts: string[] = [DEFAULT_FONT];
  const fontIds = new Map<string, number>();
  const fills: string[] = [];
  const fillIds = new Map<string, number>();
  const borders: string[] = [];
  const borderIds = new Map<string, number>();

  const formatId = (format: string | number): number => {
    if (typeof format === "number") {
      if (!Number.isInteger(format) || format < 0 || format > MAX_BUILT_IN_FORMAT) {
        throw new XlsxMiniError(
          "INVALID_NUMBER_FORMAT",
          "numberFormat",
          `Built-in number formats are the integers 0 to ${MAX_BUILT_IN_FORMAT}, not ${format}`,
        );
      }
      return format;
    }
    if (typeof format !== "string" || format.length === 0 || format.length > MAX_FORMAT_LENGTH) {
      throw new XlsxMiniError(
        "INVALID_NUMBER_FORMAT",
        "numberFormat",
        `A number format string must be 1 to ${MAX_FORMAT_LENGTH} characters`,
      );
    }
    const known = customFormats.get(format);
    if (known !== undefined) return known;
    const id = FIRST_CUSTOM_FORMAT + customFormats.size;
    customFormats.set(format, id);
    return id;
  };

  const fontSize = (size: number | undefined): number => {
    if (size === undefined) return DEFAULT_FONT_SIZE;
    if (typeof size !== "number" || !(size >= MIN_FONT_SIZE) || size > MAX_FONT_SIZE) {
      throw invalid(
        "fontSize",
        `Font sizes run from ${MIN_FONT_SIZE} to ${MAX_FONT_SIZE} points, not ${describeValue(size)}`,
      );
    }
    return size;
  };

  const fontId = (
    bold: boolean,
    italic: boolean,
    size: number | undefined,
    color: string | undefined,
  ): number => {
    const points = fontSize(size);
    const rgb = color === undefined ? "" : argb("fontColor", color);
    if (!bold && !italic && points === DEFAULT_FONT_SIZE && rgb.length === 0) return 0;
    const key = `${bold ? "b" : ""}${italic ? "i" : ""}|${points}|${rgb}`;
    const known = fontIds.get(key);
    if (known !== undefined) return known;
    fonts.push(
      "<font>" +
        (bold ? "<b/>" : "") +
        (italic ? "<i/>" : "") +
        `<sz val="${points}"/>` +
        (rgb.length === 0 ? "" : `<color rgb="${rgb}"/>`) +
        '<name val="Calibri"/><family val="2"/></font>',
    );
    const id = fonts.length - 1;
    fontIds.set(key, id);
    return id;
  };

  const fillId = (color: string | undefined): number => {
    if (color === undefined) return 0;
    const rgb = argb("fill", color);
    const known = fillIds.get(rgb);
    if (known !== undefined) return known;
    fills.push(
      `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`,
    );
    const id = fills.length + 1;
    fillIds.set(rgb, id);
    return id;
  };

  const borderId = (border: BorderStyle | Borders | undefined): number => {
    if (border === undefined) return 0;
    let sides: Sides;
    if (typeof border === "string") {
      const all = borderSide("border", border);
      sides = { top: all, right: all, bottom: all, left: all };
    } else if (typeof border === "object") {
      sides = {
        top: borderSide("border.top", border.top),
        right: borderSide("border.right", border.right),
        bottom: borderSide("border.bottom", border.bottom),
        left: borderSide("border.left", border.left),
      };
    } else {
      throw invalid("border", "A border is a line style or an object with a style per side");
    }
    const key = SIDES.map((side) => sides[side] ?? "").join("|");
    if (key === "|||") return 0;
    const known = borderIds.get(key);
    if (known !== undefined) return known;
    let xml = "<border>";
    for (const side of SIDES) {
      const style = sides[side];
      xml +=
        style === undefined
          ? `<${side}/>`
          : `<${side} style="${style}"><color auto="1"/></${side}>`;
    }
    borders.push(`${xml}<diagonal/></border>`);
    const id = borders.length;
    borderIds.set(key, id);
    return id;
  };

  const alignment = (align: string | undefined): string => {
    if (align === undefined) return "";
    if (!HORIZONTAL.has(align)) {
      throw invalid("align", `Alignments are left, center or right, not ${describeValue(align)}`);
    }
    return align;
  };

  const verticalAlignment = (align: string | undefined): string => {
    if (align === undefined) return "";
    if (!VERTICAL.has(align)) {
      throw invalid(
        "verticalAlign",
        `Vertical alignments are top, middle or bottom, not ${describeValue(align)}`,
      );
    }
    return align === "middle" ? "center" : align;
  };

  const indentation = (indent: number | undefined): number => {
    if (indent === undefined) return 0;
    if (!Number.isInteger(indent) || indent < 0 || indent > MAX_INDENT) {
      throw invalid(
        "indent",
        `Indents are whole numbers from 0 to ${MAX_INDENT}, not ${describeValue(indent)}`,
      );
    }
    return indent;
  };

  /** Excel indents only left- or right-aligned text: an indent alone aligns left, with center it is an error. */
  const indentedAlignment = (align: string, indent: number): string => {
    if (indent === 0) return align;
    if (align === "center") throw invalid("indent", "An indent cannot go with center alignment");
    return align.length === 0 ? "left" : align;
  };

  const register = (r: Resolved): number => {
    const plain =
      r.fontId === 0 &&
      r.fillId === 0 &&
      r.borderId === 0 &&
      r.align.length === 0 &&
      r.verticalAlign.length === 0 &&
      !r.wrap &&
      r.indent === 0;
    if (plain && r.numberFormat === GENERAL) return 0;
    const key = plain
      ? `${r.numberFormat}`
      : `${r.numberFormat}|${r.fontId}|${r.fillId}|${r.borderId}|${r.align}|${r.verticalAlign}|${r.wrap ? 1 : 0}|${r.indent}`;
    const known = byKey.get(key);
    if (known !== undefined) return known;
    const aligned = r.align.length > 0 || r.verticalAlign.length > 0 || r.wrap || r.indent > 0;
    const applies =
      (r.numberFormat === GENERAL ? "" : ' applyNumberFormat="1"') +
      (r.fontId === 0 ? "" : ' applyFont="1"') +
      (r.fillId === 0 ? "" : ' applyFill="1"') +
      (r.borderId === 0 ? "" : ' applyBorder="1"') +
      (aligned ? ' applyAlignment="1"' : "");
    let xml = `<xf numFmtId="${r.numberFormat}" fontId="${r.fontId}" fillId="${r.fillId}" borderId="${r.borderId}" xfId="0"${applies}`;
    if (aligned) {
      xml +=
        "><alignment" +
        (r.align.length === 0 ? "" : ` horizontal="${r.align}"`) +
        (r.verticalAlign.length === 0 ? "" : ` vertical="${r.verticalAlign}"`) +
        (r.wrap ? ' wrapText="1"' : "") +
        (r.indent === 0 ? "" : ` indent="${r.indent}"`) +
        "/></xf>";
    } else {
      xml += "/>";
    }
    cellFormats.push(xml);
    const index = cellFormats.length - 1;
    byKey.set(key, index);
    return index;
  };

  const dateOnly = (numberFormat: number): Resolved => ({
    numberFormat,
    fontId: 0,
    fillId: 0,
    borderId: 0,
    align: "",
    verticalAlign: "",
    wrap: false,
    indent: 0,
  });

  const merge = (
    cell: CellStyle | undefined,
    column: CellStyle | undefined,
    dateDefault: number,
  ): number => {
    const format = cell?.numberFormat ?? column?.numberFormat;
    const indent = indentation(cell?.indent ?? column?.indent);
    return register({
      numberFormat: format === undefined ? dateDefault : formatId(format),
      fontId: fontId(
        flag("bold", cell?.bold ?? column?.bold),
        flag("italic", cell?.italic ?? column?.italic),
        cell?.fontSize ?? column?.fontSize,
        cell?.fontColor ?? column?.fontColor,
      ),
      fillId: fillId(cell?.fill ?? column?.fill),
      borderId: borderId(cell?.border ?? column?.border),
      align: indentedAlignment(alignment(cell?.align ?? column?.align), indent),
      verticalAlign: verticalAlignment(cell?.verticalAlign ?? column?.verticalAlign),
      wrap: flag("wrap", cell?.wrap ?? column?.wrap),
      indent,
    });
  };

  return {
    indexFor(cell, column, dateDefault) {
      if (cell === undefined && column === undefined) {
        if (dateDefault === GENERAL) return 0;
        return (byDateDefault[dateDefault] ??= register(dateOnly(dateDefault)));
      }
      // Column styles were validated when the columns were collected; a cell style is met here first.
      if (cell !== undefined) styleObject(cell);
      if (cell === undefined || column === undefined) {
        const single = (cell ?? column)!;
        if (dateDefault === GENERAL) {
          const hit = bySingleObject.get(single);
          if (hit !== undefined) return hit;
          const index = merge(single, undefined, GENERAL);
          bySingleObject.set(single, index);
          return index;
        }
        let dated = bySingleDated.get(single);
        if (dated === undefined) bySingleDated.set(single, (dated = []));
        return (dated[dateDefault] ??= merge(single, undefined, dateDefault));
      }
      let columns = byPair.get(cell);
      if (columns === undefined) byPair.set(cell, (columns = new WeakMap()));
      let dated = columns.get(column);
      if (dated === undefined) columns.set(column, (dated = []));
      return (dated[dateDefault] ??= merge(cell, column, dateDefault));
    },
    validate(style) {
      styleObject(style);
      if (style.numberFormat !== undefined) formatId(style.numberFormat);
      flag("bold", style.bold);
      flag("italic", style.italic);
      fontSize(style.fontSize);
      if (style.fontColor !== undefined) argb("fontColor", style.fontColor);
      if (style.fill !== undefined) argb("fill", style.fill);
      if (style.border !== undefined && style.border !== null) {
        if (typeof style.border === "string") borderSide("border", style.border);
        else if (typeof style.border === "object") {
          for (const side of SIDES) borderSide(`border.${side}`, style.border[side]);
        } else
          throw invalid("border", "A border is a line style or an object with a style per side");
      }
      flag("wrap", style.wrap);
      verticalAlignment(style.verticalAlign);
      indentedAlignment(alignment(style.align), indentation(style.indent));
    },
    render() {
      let numberFormats = "";
      if (customFormats.size > 0) {
        numberFormats = `<numFmts count="${customFormats.size}">`;
        for (const [code, id] of customFormats)
          numberFormats += `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`;
        numberFormats += "</numFmts>";
      }
      return (
        `${XML_DECLARATION}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        numberFormats +
        `<fonts count="${fonts.length}">${fonts.join("")}</fonts>` +
        `<fills count="${fills.length + 2}">${DEFAULT_FILLS}${fills.join("")}</fills>` +
        `<borders count="${borders.length + 1}">${DEFAULT_BORDER}${borders.join("")}</borders>` +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        `<cellXfs count="${cellFormats.length}">${cellFormats.join("")}</cellXfs>` +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        "</styleSheet>"
      );
    },
  };
}
