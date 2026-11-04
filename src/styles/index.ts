import { XlsxMiniError } from "../errors/index.js";
import { XML_DECLARATION, escapeXml } from "../xml/index.js";

/**
 * How a cell looks. Set on a cell object or as a column's default; a cell's own properties win.
 *
 * @example
 * const money: CellStyle = { numberFormat: '#,##0.00 "kr"' };
 * const header: CellStyle = { bold: true };
 */
export interface CellStyle {
  readonly bold?: boolean;
  /** An Excel format string such as `"0.0%"` or `"yyyy-mm-dd"`, or a built-in format id from 0 to 49. */
  readonly numberFormat?: string | number;
}

/** Hands out cell-format indexes, deduplicated, and renders `xl/styles.xml` with exactly what was used. */
export interface StyleTable {
  /**
   * The index of the cell format for a cell style merged over a column style. `dateDefault` is the
   * built-in number format a date-like value falls back to when neither style sets one, or 0.
   */
  indexFor(cell: CellStyle | undefined, column: CellStyle | undefined, dateDefault: number): number;
  /** Throws `INVALID_NUMBER_FORMAT` for a format Excel would refuse, without registering anything. */
  validate(style: CellStyle): void;
  /** Renders the styles part. Call once, after every sheet has been rendered. */
  render(): string;
}

const GENERAL = 0;
const MAX_BUILT_IN_FORMAT = 49;
const FIRST_CUSTOM_FORMAT = 164;
const MAX_FORMAT_LENGTH = 255;

/**
 * Index 0 is always the default format. Fonts, custom number formats and cell formats are each
 * registered once per distinct value. A style object seen again on its own is resolved through a
 * WeakMap, so a column or header style reused across thousands of cells costs one lookup per cell.
 * Excel expects fill 0 to be `none`, fill 1 to be `gray125`, and a `Normal` cell style.
 */
export function createStyleTable(): StyleTable {
  const cellFormats: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const byKey = new Map<string, number>();
  const bySingleObject = new WeakMap<CellStyle, number>();
  const customFormats = new Map<string, number>();
  let boldFont = 0;

  const formatId = (format: string | number): number => {
    if (typeof format === "number") {
      if (!Number.isInteger(format) || format < 0 || format > MAX_BUILT_IN_FORMAT) {
        throw new XlsxMiniError(
          "INVALID_NUMBER_FORMAT",
          "",
          `Built-in number formats are the integers 0 to ${MAX_BUILT_IN_FORMAT}, not ${format}`,
        );
      }
      return format;
    }
    if (format.length === 0 || format.length > MAX_FORMAT_LENGTH) {
      throw new XlsxMiniError(
        "INVALID_NUMBER_FORMAT",
        "",
        `A number format string must be 1 to ${MAX_FORMAT_LENGTH} characters`,
      );
    }
    const known = customFormats.get(format);
    if (known !== undefined) return known;
    const id = FIRST_CUSTOM_FORMAT + customFormats.size;
    customFormats.set(format, id);
    return id;
  };

  const resolve = (bold: boolean, numberFormat: number): number => {
    if (!bold && numberFormat === GENERAL) return 0;
    const key = bold ? `b${numberFormat}` : `${numberFormat}`;
    const known = byKey.get(key);
    if (known !== undefined) return known;
    if (bold && boldFont === 0) boldFont = 1;
    const fontId = bold ? boldFont : 0;
    const applies =
      (numberFormat === GENERAL ? "" : ' applyNumberFormat="1"') + (bold ? ' applyFont="1"' : "");
    cellFormats.push(
      `<xf numFmtId="${numberFormat}" fontId="${fontId}" fillId="0" borderId="0" xfId="0"${applies}/>`,
    );
    const index = cellFormats.length - 1;
    byKey.set(key, index);
    return index;
  };

  const merge = (
    cell: CellStyle | undefined,
    column: CellStyle | undefined,
    dateDefault: number,
  ): number => {
    const bold = cell?.bold ?? column?.bold ?? false;
    const format = cell?.numberFormat ?? column?.numberFormat;
    return resolve(bold, format === undefined ? dateDefault : formatId(format));
  };

  return {
    indexFor(cell, column, dateDefault) {
      if (cell === undefined && column === undefined) {
        return dateDefault === GENERAL ? 0 : resolve(false, dateDefault);
      }
      if (dateDefault === GENERAL) {
        const single = column === undefined ? cell : cell === undefined ? column : undefined;
        if (single !== undefined) {
          const hit = bySingleObject.get(single);
          if (hit !== undefined) return hit;
          const index = merge(single, undefined, GENERAL);
          bySingleObject.set(single, index);
          return index;
        }
      }
      return merge(cell, column, dateDefault);
    },
    validate(style) {
      if (style.numberFormat !== undefined) formatId(style.numberFormat);
    },
    render() {
      let numberFormats = "";
      if (customFormats.size > 0) {
        numberFormats = `<numFmts count="${customFormats.size}">`;
        for (const [code, id] of customFormats)
          numberFormats += `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`;
        numberFormats += "</numFmts>";
      }
      const fonts =
        boldFont === 0
          ? '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
          : '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>';
      return (
        `${XML_DECLARATION}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        numberFormats +
        fonts +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        `<cellXfs count="${cellFormats.length}">${cellFormats.join("")}</cellXfs>` +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
        "</styleSheet>"
      );
    },
  };
}
