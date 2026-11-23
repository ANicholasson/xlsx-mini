import { describe, expect, it } from "vite-plus/test";
import { XlsxReadError } from "../../errors/read.js";
import { type ByteChunks, readXml } from "../../xml/read.js";
import { type DateStyles, stylesVisitor } from "../read.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const encoder = new TextEncoder();

function bytesOf(part: string): ByteChunks {
  let given = false;
  return {
    read() {
      const chunk = given ? undefined : encoder.encode(part);
      given = true;
      return Promise.resolve(chunk);
    },
    close() {},
  };
}

const escaped = (text: string) =>
  text.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;");
const numFmt = (id: number | string, code: string) =>
  `<numFmt numFmtId="${id}" formatCode="${escaped(code)}"/>`;
const xf = (id?: number | string, more = ' xfId="0"') =>
  `<xf${id === undefined ? "" : ` numFmtId="${id}"`} fontId="0" fillId="0" borderId="0"${more}/>`;
const styleSheet = (inner: string) =>
  `<styleSheet xmlns="${MAIN}" xmlns:mc="${COMPATIBILITY}">${inner}</styleSheet>`;
const alternate = (inner: string) =>
  `<mc:AlternateContent><mc:Choice Requires="x14">${inner}</mc:Choice></mc:AlternateContent>`;
const named = '<cellStyleXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellStyleXfs>';

async function datesOf(part: string): Promise<readonly number[]> {
  const dates: DateStyles = await readXml(bytesOf(part), stylesVisitor());
  return [...dates];
}

async function outcomeOf(part: string): Promise<readonly number[] | string> {
  try {
    return await datesOf(part);
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    return error.code;
  }
}

/** Whether each code shows a date, as the one style of a part that names it. */
async function dateCodes(codes: readonly string[]): Promise<readonly boolean[]> {
  const part = styleSheet(
    `<numFmts>${codes.map((code, at) => numFmt(164 + at, code)).join("")}</numFmts>` +
      `<cellXfs>${codes.map((_, at) => xf(164 + at)).join("")}</cellXfs>`,
  );
  return (await datesOf(part)).map((date) => date === 1);
}

/**
 * Number format codes and whether a number they format reads as a date. Where the answer comes from:
 * ECMA-376 part 1, 18.8.31, or the corpus file in which an application wrote the code, read back by
 * LibreOffice in `reader-proto/expected`. Every code not from the corpus was also read in one
 * hand-made workbook by LibreOffice 26.8 and openpyxl 3.1.5; where they disagree, the reader that
 * agrees is the source, and the fourth column says what the other shows. A time of day and an
 * elapsed time read as numbers by ADR 0032, where LibreOffice shows a time and openpyxl a time or
 * a duration; that is not counted as a disagreement.
 */
const CODES: readonly (readonly [code: string, date: boolean, source: string, other?: string])[] = [
  // Quoted text, an escaped character, and the character after an underscore or an asterisk.
  ['0 "d"', false, "18.8.31"],
  ['"yyyy" 0', false, "18.8.31"],
  ['"Date: "yyyy-mm-dd', true, "18.8.31"],
  ['"Week of" d mmmm', true, "numbers/type-matrix.xlsx"],
  ['"Week of "d\\ mmmm', true, "libreoffice/type-matrix.xlsx"],
  ['"Week of"\\ d\\ mmmm', true, "excel/type-matrix.xlsx"],
  ['"day" 0', false, "xlsx-mini/type-matrix.xlsx"],
  ['"day"\\ 0', false, "excel/type-matrix.xlsx"],
  ['0"d"', false, "numbers/type-matrix.xlsx"],
  ["0\\d", false, "excel/type-matrix.xlsx"],
  ["\\y\\y\\y\\y", false, "18.8.31"],
  ["0_d", false, "18.8.31"],
  ["_y0", false, "18.8.31"],
  ["*d0", false, "18.8.31", "openpyxl: a date"],
  ["0*m", false, "18.8.31", "openpyxl: a date"],
  ['0"d', false, "LibreOffice", "openpyxl: a date; the quote is not closed"],
  ['yyyy"', true, "18.8.31"],
  ["0\\", false, "18.8.31"],
  ["0_", false, "18.8.31"],
  ["0*", false, "18.8.31"],
  // Bracketed sections: colours, locales and currencies, conditions.
  ["[Red]0", false, "18.8.31"],
  ["[Red]yyyy-mm-dd", true, "18.8.31"],
  ["[Red]0.00;[Blue]-0.00", false, "xlsx-mini/type-matrix.xlsx"],
  ["[RED]0.00;[BLUE]\\-0.00", false, "libreoffice/type-matrix.xlsx"],
  ["[Blue]d", true, "18.8.31"],
  ["[Color10]0", false, "18.8.31"],
  ["[$-409]0", false, "18.8.31"],
  ["[$-409]d-mmm-yy;@", true, "xlsx-mini/type-matrix.xlsx"],
  ["[$-409]d\\-mmm\\-yy;@", true, "excel/type-matrix.xlsx"],
  ["[$€-2] #,##0.00", false, "18.8.31"],
  ["[$d-409]0", false, "18.8.31"],
  ["[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy", true, "18.8.31"],
  ["[$-x-sysdate]dddd, mmmm dd, yyyy", true, "openpyxl", "LibreOffice: does not parse it"],
  ["[$-F400]h:mm:ss\\ AM/PM", false, "18.8.31"],
  ["[>100]0;0", false, "18.8.31"],
  ["[>100]yyyy;0", true, "openpyxl", "LibreOffice: a number shown as a year"],
  ["[<1]h:mm;yyyy-mm-dd", false, "LibreOffice"],
  ["[Red 0", false, "LibreOffice", "openpyxl: a date, from the d of Red"],
  // Elapsed time: a time, which stays a number, unless a date stands beside it.
  ["[h]:mm", false, "excel/type-matrix.xlsx"],
  ["[h]:mm:ss", false, "18.8.31"],
  ["[mm]:ss", false, "18.8.31"],
  ["[ss].00", false, "18.8.31"],
  ["mm:[ss]", false, "18.8.31"],
  ["[h]", false, "18.8.31"],
  ["[m]", false, "18.8.31"],
  ["[H]:MM", false, "18.8.31"],
  ["[h]:mm yyyy", true, "LibreOffice", "openpyxl: a duration"],
  ["[hhm]:ss", false, "LibreOffice"],
  ["[hs]:mm", true, "openpyxl", "LibreOffice: a number"],
  ["[mh]", false, "openpyxl", "LibreOffice: a date"],
  // m: a minute after an hour or before a second, a month otherwise.
  ["h:mm", false, "18.8.31"],
  ["hh:mm:ss", false, "18.8.31"],
  ["mm:ss", false, "18.8.31"],
  ["m:s", false, "18.8.31"],
  ["mm:ss.0", false, "libreoffice/type-matrix.xlsx"],
  ["mmss.0", false, "18.8.31"],
  ['h" "mm', false, "18.8.31"],
  ["h mm", false, "18.8.31"],
  ["mm ss", false, "18.8.31"],
  ["h:mm d", true, "18.8.31"],
  ["h:[Red]mm", false, "18.8.31"],
  ["mm h", true, "18.8.31"],
  ["m", true, "18.8.31"],
  ["mm", true, "18.8.31"],
  ["mmmmm", true, "18.8.31"],
  ["mm d", true, "18.8.31"],
  ["d mm", true, "18.8.31"],
  ["mm-yy", true, "18.8.31"],
  ["yyyy mm", true, "18.8.31"],
  ["mm/dd/yyyy", true, "libreoffice/cell-types.xlsx"],
  ["dd/mm/yyyy hh:mm", true, "numbers/type-matrix.xlsx"],
  ["dd/mm/yyyy\\ hh:mm", true, "excel/type-matrix.xlsx"],
  ["mm/dd/yy\\ hh:mm\\ AM/PM", true, "libreoffice/cell-types.xlsx"],
  ["yyyy-mm-dd hh:mm", true, "numbers/orders.xlsx"],
  ["yyyy\\-mm\\-dd\\ hh:mm", true, "excel/orders.xlsx"],
  ["h", false, "18.8.31"],
  ["s", false, "18.8.31"],
  ["hh:mm:ss.000", false, "18.8.31"],
  ["h:mm:ss;@", false, "18.8.31"],
  // AM/PM and A/P mark a time; their letters are not tokens.
  ["h:mm AM/PM", false, "xlsx-mini/type-matrix.xlsx"],
  ["h:mm\\ AM/PM", false, "libreoffice/type-matrix.xlsx"],
  ["h A/P", false, "18.8.31"],
  ["h a/p", false, "18.8.31"],
  ["am/pm", false, "18.8.31"],
  ["A/P", false, "18.8.31"],
  ["0 AM/PM", false, "LibreOffice"],
  ["mm AM/PM", true, "LibreOffice"],
  ["AM/PM mm", true, "LibreOffice"],
  ["h:mm AM/P", true, "LibreOffice"],
  // Sections: the first decides, since a number that can be a date is positive.
  ["0;yyyy", false, "18.8.31"],
  ["yyyy;0", true, "openpyxl", "LibreOffice: a number shown as a year"],
  ["0;0;yyyy", false, "18.8.31"],
  ["0;0;0;yyyy", false, "18.8.31"],
  ["yyyy;;", true, "18.8.31"],
  [";yyyy", false, "18.8.31"],
  ["@;yyyy", false, "18.8.31"],
  ["h:mm;yyyy", false, "LibreOffice"],
  ['0.00;"--"0.00', false, "numbers/type-matrix.xlsx"],
  // General as a custom code, which LibreOffice and Numbers write, in every case of its letters.
  ["General", false, "libreoffice/minimal.xlsx"],
  ["general", false, "18.8.31"],
  ["GENERAL", false, "18.8.31"],
  ["gENERAL", false, "18.8.31"],
  ["General;yyyy", false, "18.8.31"],
  ['"x"General', false, "18.8.31"],
  ['General "d"', false, "18.8.31"],
  ["Generally", false, "LibreOffice", "openpyxl: a date, from the y"],
  ["yyyy General", false, "LibreOffice", "openpyxl: a date"],
  ["General d", false, "LibreOffice", "openpyxl: a date"],
  // An exponent, text, nothing, and numbers.
  ["0.00E+00", false, "libreoffice/type-matrix.xlsx"],
  ["0.00e+00", false, "18.8.31"],
  ["##0.0E+0", false, "18.8.31"],
  ["0.00E-00", false, "18.8.31"],
  ["0.00E00", false, "LibreOffice"],
  ["@", false, "libreoffice/type-matrix.xlsx"],
  ['@ "d"', false, "18.8.31"],
  ["", false, "18.8.31"],
  ["# ?/?", false, "xlsx-mini/type-matrix.xlsx"],
  ["0.0%", false, "excel/styles.xlsx"],
  ['#,##0.00 "kr"', false, "xlsx-mini/appearance.xlsx"],
  ['#,##0.00" kr"', false, "libreoffice/appearance.xlsx"],
  ['#,##0.00\\ "kr"', false, "excel/appearance.xlsx"],
  ['0.00 "m"', false, "xlsx-mini/type-matrix.xlsx"],
  ['0.00\\ "m"', false, "excel/type-matrix.xlsx"],
  // Years and days, in either case.
  ["yyyy", true, "excel/type-matrix.xlsx"],
  ["yyyy-mm-dd", true, "numbers/auto-width.xlsx"],
  ["yyyy\\-mm\\-dd", true, "excel/auto-width.xlsx"],
  ["d mmm yyyy", true, "numbers/type-matrix.xlsx"],
  ["d\\ mmm\\ yyyy", true, "excel/type-matrix.xlsx"],
  ["yy", true, "18.8.31"],
  ["dddd", true, "18.8.31"],
  ["YYYY", true, "18.8.31"],
  ["DD/MM/YYYY", true, "18.8.31"],
  ["Y", true, "openpyxl", "LibreOffice: the letter Y"],
  ["S", false, "18.8.31"],
  // Era and calendar letters: a date only with y, m or d beside them.
  ["e", false, "openpyxl", "LibreOffice: a year of an era"],
  ["ee", false, "openpyxl", "LibreOffice: a year of an era"],
  ["g", false, "openpyxl", "LibreOffice: an era"],
  ["ggg", false, "openpyxl", "LibreOffice: an era"],
  ["[$-411]e", false, "openpyxl", "LibreOffice: a year of an era"],
  ['e"年"', false, "openpyxl", "LibreOffice: a year of an era"],
  ["ge.m.d", true, "18.8.31"],
  ['[$-411]ggge"年"m"月"d"日"', true, "18.8.31"],
  ["b", false, "LibreOffice"],
  ["bbbb", false, "LibreOffice"],
  ["[$-41E]bbbb/mm/dd", true, "openpyxl", "LibreOffice: does not parse it"],
  ["B2yyyy", true, "18.8.31"],
  ["aaa", false, "openpyxl", "LibreOffice: a day of the week"],
  ["nn", false, "openpyxl", "LibreOffice: a day of the week, by a code of its own"],
  ["q", false, "openpyxl", "LibreOffice: a quarter, by a code of its own"],
  ["ww", false, "openpyxl", "LibreOffice: a week, by a code of its own"],
];

describe("which cell styles show a date", () => {
  it("reads each number format code as the standard and the applications show it, from a table of codes", async () => {
    const codes = CODES.map(([code]) => code);
    expect(new Set(codes).size).toBe(codes.length);
    const read = await dateCodes(codes);
    expect(Object.fromEntries(codes.map((code, at) => [code, read[at]]))).toEqual(
      Object.fromEntries(CODES.map(([code, date]) => [code, date])),
    );
  });

  it("reads a built-in id as ADR 0032 lists it: 14 to 17, 22, 27 to 31, 36 and 50 to 58 are dates", async () => {
    const ids = Array.from({ length: 82 }, (_, id) => id);
    const part = styleSheet(`<cellXfs>${ids.map((id) => xf(id)).join("")}</cellXfs>`);
    const dates = await datesOf(part);
    const listed = [14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58];
    expect(ids.filter((id) => dates[id] === 1)).toEqual(listed);
  });

  it("lets a code replace a built-in id, 0 included, and takes the last code of an id, as LibreOffice and openpyxl do", async () => {
    const replaced = styleSheet(
      `<numFmts>${numFmt(14, "0.00")}${numFmt(20, "yyyy-mm-dd")}${numFmt(0, "d/m")}</numFmts>` +
        `<cellXfs>${xf(0)}${xf(14)}${xf(20)}${xf(22)}</cellXfs>`,
    );
    expect(await datesOf(replaced)).toEqual([1, 0, 1, 1]);
    for (const [first, last, date] of [
      ["yyyy-mm-dd", "0.00", 0],
      ["0.00", "yyyy-mm-dd", 1],
    ] as const) {
      const twice = styleSheet(
        `<numFmts>${numFmt(164, first)}${numFmt(164, last)}</numFmts><cellXfs>${xf(0)}${xf(164)}</cellXfs>`,
      );
      expect(await datesOf(twice)).toEqual([0, date]);
    }
  });

  it("takes a style's number format from its own numFmtId alone, 0 when it has none, as LibreOffice and openpyxl do", async () => {
    const styles = [
      xf(undefined, ' xfId="1"'),
      xf(0, ' xfId="1" applyNumberFormat="0"'),
      xf(14, ' applyNumberFormat="0"'),
      xf(164, ' applyNumberFormat="false"'),
      xf("014"),
    ];
    const part = styleSheet(
      `<numFmts>${numFmt(164, "yyyy")}</numFmts>${named}<cellXfs>${styles.join("")}</cellXfs>`,
    );
    expect(await datesOf(part)).toEqual([0, 0, 1, 1, 1]);
    const general = styleSheet(`<numFmts>${numFmt(0, "yyyy")}</numFmts><cellXfs>${xf()}</cellXfs>`);
    expect(await datesOf(general)).toEqual([1]);
  });

  it("reads a built-in id the part does not define as the standard's format, or General", async () => {
    const part = styleSheet(`<cellXfs>${xf(23)}${xf(14)}${xf(80)}${xf(81)}</cellXfs>`);
    expect(await datesOf(part)).toEqual([0, 1, 0, 0]);
  });

  it("has one style that shows no date when the part has no cell styles", async () => {
    for (const inner of ["", "<cellXfs/>", `<fonts/>${named}`]) {
      expect(await datesOf(styleSheet(inner)), inner).toEqual([0]);
    }
  });

  it("passes over what holds no cell style: named styles, differential formats, extensions, and a numFmt or an xf outside its list", async () => {
    const part = styleSheet(
      `<numFmts>${numFmt(164, "0.00")}</numFmts>` +
        `<cellStyleXfs>${xf(14)}</cellStyleXfs>` +
        `<dxfs><dxf>${numFmt(164, "yyyy")}</dxf></dxfs>` +
        `<cellXfs>${xf(0)}${xf(164)}<xf numFmtId="0"><alignment/><extLst><ext/></extLst></xf></cellXfs>` +
        xf(14) +
        alternate(`<x14:dxfs><dxf>${numFmt(165, "yyyy")}</dxf></x14:dxfs>`) +
        `<extLst><ext><x14:slicerStyles/></ext></extLst>`,
    );
    expect(await datesOf(part)).toEqual([0, 0, 0]);
  });

  it("refuses a style that names an id past the built-in ones that numFmts does not define, for which LibreOffice takes a differential format's code and openpyxl reads General", async () => {
    const differential = (id: number, code: string) =>
      `<dxfs><dxf>${numFmt(id, code)}</dxf></dxfs>`;
    for (const part of [
      styleSheet(`<cellXfs>${xf(0)}${xf(165)}</cellXfs>${differential(165, "yyyy-mm-dd")}`),
      styleSheet(`${differential(82, "yyyy-mm-dd")}<cellXfs>${xf(82)}</cellXfs>`),
      styleSheet(`${numFmt(165, "0.00")}<cellXfs>${xf(165)}</cellXfs>`),
      styleSheet(`<cellXfs>${xf(100)}${xf(200)}${xf(4_294_967_310)}</cellXfs>`),
      styleSheet(`<cellXfs>${xf(82)}</cellXfs>`),
    ]) {
      expect(await outcomeOf(part), part).toBe("INVALID_WORKBOOK");
    }
    const defined = styleSheet(
      `<numFmts>${numFmt(165, "0.00")}${numFmt(82, "d")}</numFmts><cellXfs>${xf(165)}${xf(82)}</cellXfs>${differential(165, "yyyy")}`,
    );
    expect(await datesOf(defined)).toEqual([0, 1]);
    const builtIn = styleSheet(
      `<cellXfs>${xf(14)}${xf(19)}${xf(81)}</cellXfs><dxfs><dxf>${numFmt(14, "0.00")}</dxf><dxf>${numFmt(19, "yyyy/m/d")}</dxf><dxf>${numFmt(81, "yyyy")}</dxf></dxfs>`,
    );
    expect(await datesOf(builtIn)).toEqual([1, 0, 0]);
  });

  it("refuses a numFmts or a cellXfs stated twice, out of order, or anywhere but directly under the root, where LibreOffice and openpyxl read it in different ways", async () => {
    const formats = `<numFmts>${numFmt(164, "yyyy")}</numFmts>`;
    const styles = `<cellXfs>${xf(0)}${xf(14)}</cellXfs>`;
    for (const inner of [
      formats + formats + styles,
      styles + formats,
      styles + formats + styles,
      alternate(formats) + styles,
      formats + alternate(styles),
      styles + alternate(styles),
      `<extLst><ext>${formats}</ext></extLst>${styles}`,
      `<dxfs><dxf>${styles}</dxf></dxfs>`,
    ]) {
      expect(await outcomeOf(styleSheet(inner)), inner).toBe("INVALID_WORKBOOK");
    }
  });

  it("refuses anything in a numFmts but a numFmt that holds nothing, and anything in a cellXfs but an xf, or an xf deeper", async () => {
    for (const inner of [
      `<numFmts><extLst>${numFmt(164, "yyyy")}</extLst></numFmts>`,
      `<numFmts>${alternate(numFmt(164, "yyyy"))}</numFmts>`,
      `<numFmts><numFmt numFmtId="164" formatCode="yyyy"><extLst/></numFmt></numFmts>`,
      `<numFmts>${numFmt(14, "0")}<numFmtx numFmtId="14" formatCode="0"/></numFmts>`,
      `<numFmts><numFmt numFmtId="14" formatCode="0">${numFmt(15, "0")}</numFmt></numFmts>`,
      `<cellXfs>${xf(0)}<extLst>${xf(14)}</extLst></cellXfs>`,
      `<cellXfs>${xf(0)}<xf numFmtId="0">${xf(14)}</xf></cellXfs>`,
      `<cellXfs>${alternate(xf(14))}</cellXfs>`,
      `<cellXfs><cellStyle/></cellXfs>`,
    ]) {
      expect(await outcomeOf(styleSheet(inner)), inner).toBe("INVALID_WORKBOOK");
    }
  });

  it("refuses a numFmt without an id or a code, and an id that is not a decimal integer, which LibreOffice or openpyxl reads otherwise", async () => {
    const bad = [" 14", "+14", "14.0", "1.4e1", "abc", "-14", "", "&#49;4"];
    for (const id of bad) {
      const part = styleSheet(
        `<numFmts>${numFmt(id, "0.00")}</numFmts><cellXfs>${xf(14)}</cellXfs>`,
      );
      expect(await outcomeOf(part), `numFmt ${id}`).toBe("INVALID_WORKBOOK");
      expect(await outcomeOf(styleSheet(`<cellXfs>${xf(id)}</cellXfs>`)), `xf ${id}`).toBe(
        "INVALID_WORKBOOK",
      );
    }
    for (const numFmt of ['<numFmt numFmtId="14"/>', '<numFmt formatCode="0.00"/>']) {
      const part = styleSheet(`<numFmts>${numFmt}</numFmts><cellXfs>${xf(14)}</cellXfs>`);
      expect(await outcomeOf(part), numFmt).toBe("INVALID_WORKBOOK");
    }
  });

  it("refuses a root that is not styleSheet, which openpyxl reads and LibreOffice does not", async () => {
    for (const part of [
      `<styles xmlns="${MAIN}"><cellXfs>${xf(14)}</cellXfs></styles>`,
      "<sst/>",
    ]) {
      expect(await outcomeOf(part), part).toBe("INVALID_WORKBOOK");
    }
  });

  it("reads each code once, however many styles name it", async () => {
    const code = `${"0".repeat(1_000_000)} yyyy`;
    const part = styleSheet(
      `<numFmts>${numFmt(164, code)}</numFmts><cellXfs>${xf(164).repeat(200_000)}</cellXfs>`,
    );
    const started = performance.now();
    const dates = await datesOf(part);
    expect(performance.now() - started).toBeLessThan(2_000);
    expect(dates).toHaveLength(200_000);
    expect(dates.every((date) => date === 1)).toBe(true);
  });
});
