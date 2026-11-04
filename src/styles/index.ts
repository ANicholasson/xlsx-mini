import { XML_DECLARATION } from "../xml/index.js";

const BUILT_IN_DATE = 14;
const BUILT_IN_DATE_TIME = 22;

/** Hands out cell-format indexes on first use and renders `xl/styles.xml` with exactly those. */
export interface StyleTable {
  dateOnly(): number;
  dateTime(): number;
  /** Renders the styles part. Call once, after every sheet has been rendered. */
  render(): string;
}

/**
 * The styles part always carries the default cell format at index 0 and, when a sheet asked for
 * them, one format per built-in date layout. Excel expects fill 0 to be `none`, fill 1 to be
 * `gray125`, and a `Normal` cell style.
 */
export function createStyleTable(): StyleTable {
  const cellFormats: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  let dateOnly = 0;
  let dateTime = 0;

  const add = (numberFormat: number): number => {
    cellFormats.push(
      `<xf numFmtId="${numberFormat}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`,
    );
    return cellFormats.length - 1;
  };

  return {
    dateOnly() {
      if (dateOnly === 0) dateOnly = add(BUILT_IN_DATE);
      return dateOnly;
    },
    dateTime() {
      if (dateTime === 0) dateTime = add(BUILT_IN_DATE_TIME);
      return dateTime;
    },
    render() {
      return (
        `${XML_DECLARATION}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
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
