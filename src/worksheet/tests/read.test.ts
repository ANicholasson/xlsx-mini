import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { type ReadLimits, XlsxReadError, resolveLimits } from "../../errors/read.js";
import { type ByteChunks, scanXml } from "../../xml/read.js";
import { serialFromEpochMillis } from "../index.js";
import { type ReadValue, worksheetVisitor } from "../read.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const encoder = new TextEncoder();

function bytesOf(...chunks: string[]): ByteChunks {
  let next = 0;
  return {
    read: () => Promise.resolve(next < chunks.length ? encoder.encode(chunks[next++]) : undefined),
    close() {},
  };
}

const worksheet = (inner: string) => `<worksheet xmlns="${MAIN}">${inner}</worksheet>`;
const data = (rows: string) => worksheet(`<sheetData>${rows}</sheetData>`);
const alternate = (inner: string) =>
  `<mc:AlternateContent xmlns:mc="${COMPATIBILITY}"><mc:Choice Requires="x14">${inner}</mc:Choice></mc:AlternateContent>`;
const cell = (attributes: string, inner: string) =>
  data(`<row r="1"><c r="A1"${attributes}>${inner}</c></row>`);

type Rows = readonly (readonly [number, ...ReadValue[]])[];

interface Reading {
  readonly strings?: readonly string[];
  readonly limits?: ReadLimits;
  /** Which styles show a date, by index: 1 for a date style. None by default. */
  readonly dates?: readonly number[];
  readonly date1904?: boolean;
}

async function rowsOf(
  part: string,
  { strings = [], limits, dates, date1904 = false }: Reading = {},
): Promise<Rows> {
  const { positions, rowText } = resolveLimits(limits).rows();
  const rows: (readonly [number, ...ReadValue[]])[] = [];
  const visitor = worksheetVisitor({
    strings,
    dates: dates && Uint8Array.from(dates),
    date1904,
    positions,
    rowText,
  });
  for await (const { number, cells } of scanXml(bytesOf(part), visitor)) {
    rows.push([number, ...cells]);
  }
  return rows;
}

async function outcomeOf(part: string, reading?: Reading): Promise<Rows | string> {
  try {
    return await rowsOf(part, reading);
  } catch (error) {
    if (!(error instanceof XlsxReadError)) throw error;
    return [error.code, error.cell ?? error.limit].filter(Boolean).join(" ");
  }
}

describe("the rows of a worksheet part", () => {
  it("gives each row that holds a value, with its number and its cells from column A, as LibreOffice 26.8 reads them", async () => {
    const sparse = data(
      '<row r="3"><c r="C3"><v>1</v></c></row><row r="4" spans="1:9"/>' +
        '<row r="7"><c r="A7" t="inlineStr"><is><t/></is></c><c r="B7"><v>2</v></c><c r="D7" s="1"/></row>',
    );
    expect(await rowsOf(sparse, { dates: [0, 1] })).toEqual([
      [3, null, null, 1],
      [7, null, 2],
    ]);
    const unreferenced = data(
      '<row><c><v>1</v></c><c><v>2</v></c></row><row><c r="B2"><v>3</v></c><c><v>4</v></c></row>',
    );
    expect(await rowsOf(unreferenced)).toEqual([
      [1, 1, 2],
      [2, null, 3, 4],
    ]);
    const corner = await rowsOf(data('<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>'));
    expect(corner.map(([number, ...cells]) => [number, cells.length, cells.at(-1)])).toEqual([
      [1_048_576, 16_384, 1],
    ]);
    expect(corner[0]!.slice(1, -1).every((value) => value === null)).toBe(true);
  });

  it("reads every type of cell as the file stores it", async () => {
    const strings = ["shared", "", "_x0041_ as decoded by the table"];
    const cells = [
      "<v>1.5</v>",
      '<c t="n"><v>-2</v></c>',
      '<c t="s"><v>0</v></c>',
      '<c t="s"><v>02</v></c>',
      '<c t="str"><f>A1&amp;""</f><v>__x0078_0041_ and _x0041_</v></c>',
      '<c t="inlineStr"><is><r><rPr><b/></rPr><t>in</t></r><r><t xml:space="preserve">line _x0041_ </t></r><rPh sb="0" eb="1"><t>P</t></rPh></is></c>',
      '<c t="b"><v>1</v></c>',
      '<c t="b"><v>0</v></c>',
      '<c t="e"><v>#DIV/0!</v></c>',
      '<c t="str"><v>a&#13;b\r\nc</v></c>',
      "<f>SUM(A1:A2)</f><v>7</v>",
    ].map((inner, at) => {
      const reference = `${String.fromCharCode(65 + at)}1`;
      return inner.startsWith("<c")
        ? inner.replace("<c", `<c r="${reference}"`)
        : `<c r="${reference}">${inner}</c>`;
    });
    expect(await rowsOf(data(`<row r="1">${cells.join("")}</row>`), { strings })).toEqual([
      [
        1,
        1.5,
        -2,
        "shared",
        "_x0041_ as decoded by the table",
        "_x0041_ and A",
        "inline A ",
        true,
        false,
        "#DIV/0!",
        "a\rb\nc",
        7,
      ],
    ]);
  });

  it("reads a cell that stores nothing as empty, whatever its type, and an empty v as empty where the type is text or a date", async () => {
    const stored = (type: string, inner: string) =>
      data(`<row r="1"><c r="A1"${type}>${inner}</c><c r="B1"><v>2</v></c></row>`);
    const types = [
      "",
      ' t="n"',
      ' t="s"',
      ' t="b"',
      ' t="e"',
      ' t="str"',
      ' t="d"',
      ' t="inlineStr"',
    ];
    for (const type of types) {
      for (const inner of ["<f>1+1</f>", ""]) {
        expect(await rowsOf(stored(type, inner)), `${type} ${inner}`).toEqual([[1, null, 2]]);
      }
    }
    for (const type of [' t="e"', ' t="str"', ' t="d"']) {
      for (const inner of ["<v></v>", "<v/>", '<f>""</f><v xml:space="preserve"></v>']) {
        expect(await rowsOf(stored(type, inner)), `${type} ${inner}`).toEqual([[1, null, 2]]);
      }
    }
    for (const type of ["", ' t="n"', ' t="s"', ' t="b"']) {
      for (const inner of ["<v></v>", "<v/>"]) {
        const part = stored(type, inner);
        expect(await outcomeOf(part, { strings: [""] }), `${type} ${inner}`).toBe(
          "INVALID_CELL A1",
        );
      }
    }
    const blank = data(
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is/></c></row><row r="2"><c r="A2" t="inlineStr"><is><t></t></is></c></row>',
    );
    expect(await rowsOf(blank, { strings: [""] })).toEqual([]);
  });

  it("reads a number as XML Schema's double spells one, white space around it included", async () => {
    const numbers: Readonly<Record<string, number>> = {
      " 12 ": 12,
      "+5": 5,
      "1e3": 1000,
      "1E3": 1000,
      ".5": 0.5,
      "5.": 5,
      "-0": -0,
      "007": 7,
      "\t3\n": 3,
      "1e+2": 100,
      "1.5e-3": 0.0015,
      "-.25E-1": -0.025,
      "&#49;": 1,
      "1e-400": 0,
      "179769313486231570000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000":
        Number.MAX_VALUE,
    };
    for (const [text, value] of Object.entries(numbers)) {
      expect(await rowsOf(cell("", `<v>${text}</v>`)), text).toEqual([[1, value]]);
    }
  });

  it("refuses with INVALID_CELL a number that is not decimal or not finite, which LibreOffice reads as something else", async () => {
    const texts = [
      "INF",
      "-INF",
      "NaN",
      "Infinity",
      "1e999",
      "-1e999",
      "0x10",
      "0b1",
      "1,5",
      "1 2",
      "1e",
      "1e+",
      "e3",
      "1_0",
      "1.2.3",
      "--1",
      "+-1",
      "١٢",
      ".",
      "+",
      " ",
      "1e3.5",
    ];
    for (const text of texts) {
      expect(await outcomeOf(cell("", `<v>${text}</v>`)), text).toBe("INVALID_CELL A1");
    }
  });

  it("refuses a boolean other than 0 or 1, and an index that names no shared string", async () => {
    for (const text of ["true", "TRUE", "false", "2", "-1", " 1", "01"]) {
      expect(await outcomeOf(cell(' t="b"', `<v>${text}</v>`)), text).toBe("INVALID_CELL A1");
    }
    const strings = ["zero", "one"];
    for (const text of ["-1", " 1", "1.0", "1e0", "2", "99", "length", "constructor"]) {
      expect(await outcomeOf(cell(' t="s"', `<v>${text}</v>`), { strings }), text).toBe(
        "INVALID_CELL A1",
      );
    }
    expect(await rowsOf(cell(' t="s"', "<v>01</v>"), { strings })).toEqual([[1, "one"]]);
  });

  it("refuses an unknown type", async () => {
    const types = [
      "x",
      "",
      "S",
      "N",
      "inlinestr",
      "inlineStx",
      "inlineString",
      "string",
      "&#110;",
      "nn",
      "str ",
    ];
    for (const type of types) {
      for (const inner of ["<v>1</v>", "", "<f>1</f>"]) {
        const part = cell(` t="${type}"`, inner);
        await expect(rowsOf(part), `${type} ${inner}`).rejects.toMatchObject({
          code: "INVALID_CELL",
          cell: "A1",
          message: "a cell is of an unknown type",
        });
      }
    }
  });

  it("refuses a row out of order, numbered 0, past the last row or not in decimal, naming its first cell", async () => {
    const one = (number: string) => `<row r="${number}"><c><v>1</v></c></row>`;
    const rows = {
      [one("0")]: "INVALID_CELL",
      [one("2") + one("1")]: "INVALID_CELL A1",
      [one("2") + one("2")]: "INVALID_CELL A2",
      [one("1048577")]: "INVALID_CELL A1048577",
      [one("1048576") + "<row/>"]: "INVALID_CELL A1048577",
      [one("+1")]: "INVALID_CELL",
      [one(" 1")]: "INVALID_CELL",
      [one("1.0")]: "INVALID_CELL",
      [one("")]: "INVALID_CELL",
    };
    for (const [rowsXml, expected] of Object.entries(rows)) {
      expect(await outcomeOf(data(rowsXml)), rowsXml).toBe(expected);
    }
  });

  it("refuses a cell reference that is malformed, names another row or is not right of the cell before it, as the file writes it", async () => {
    const references = [
      "a1",
      "$A$1",
      "A01",
      "AAAA1",
      "XFE1",
      "A2",
      "A0",
      "1",
      "A",
      "A1 ",
      " A1",
      "",
      "A&#49;",
      "AAAAAAAAA1",
      "A\t1",
    ];
    for (const reference of references) {
      const part = data(`<row r="1"><c r="${reference}"><v>1</v></c></row>`);
      expect(await outcomeOf(part), reference).toBe(
        ["INVALID_CELL", reference].filter(Boolean).join(" "),
      );
    }
    const long = {
      AAAAAAAAAA1: "AAAAAAAAAA…",
      [`${"\t".repeat(4_194_000)}A1`]: `${"\t".repeat(10)}…`,
      [`A${"&#49;".repeat(20)}`]: "A&#49;&#49…",
      "AAAAAAAAA\u{1F600}1": "AAAAAAAAA…",
      "AAAAAAAA\u{1F600}1": "AAAAAAAA\u{1F600}…",
      "AAAAAAAAA\u{10FFFF}1": "AAAAAAAAA…",
    };
    for (const [reference, written] of Object.entries(long)) {
      const part = data(`<row r="1"><c r="${reference}"><v>1</v></c></row>`);
      expect(await outcomeOf(part), reference.slice(0, 20)).toBe(`INVALID_CELL ${written}`);
    }
    const order = {
      '<c r="B1"><v>1</v></c><c r="A1"><v>2</v></c>': "INVALID_CELL A1",
      '<c r="A1"><v>1</v></c><c r="A1"><v>2</v></c>': "INVALID_CELL A1",
      '<c r="XFD1"><v>1</v></c><c><v>2</v></c>': "INVALID_CELL XFE1",
    };
    for (const [cells, expected] of Object.entries(order)) {
      expect(await outcomeOf(data(`<row r="1">${cells}</row>`)), cells).toBe(expected);
    }
    const unnumbered = data(
      '<row r="2"><c r="A2"><v>1</v></c></row><row><c r="A5"><v>2</v></c></row>',
    );
    expect(await outcomeOf(unnumbered)).toBe("INVALID_CELL A5");
  });

  it("refuses what a row or a cell cannot hold, naming the cell it stands in or the position after the last", async () => {
    const rows = {
      '<row r="1"><c r="A1"><v>1</v></c><foo/></row>': "INVALID_CELL B1",
      '<row r="1"><foo><c r="A1"><v>1</v></c></foo></row>': "INVALID_CELL A1",
      [`<row r="1">${alternate('<c r="A1"><v>1</v></c>')}</row>`]: "INVALID_CELL A1",
      '<row r="1"><c r="A1"><foo/><v>1</v></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1"><foo><v>1</v></foo></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1"><v>1</v><v>2</v></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1"><v>1<b/>2</v></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1"><v><t>5</t></v></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="str"><v><t>5</t></v></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1"><x>5</x></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="inlineStr"><v>text</v></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1"><is><t>text</t></is></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="s"><is><t>text</t></is></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="str"><is><t>text</t></is></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="inlineStr"><is><t>text</t></is><v>v</v></c></row>':
        "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is><is><t>b</t></is></c></row>':
        "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="inlineStr"><is><foo/><t>a</t></is></c></row>': "INVALID_CELL A1",
      '<row r="1"><c r="A1" t="inlineStr"><is><t>a</t><t>b</t></is></c></row>': "INVALID_CELL A1",
    };
    for (const [rowsXml, expected] of Object.entries(rows)) {
      expect(await outcomeOf(data(rowsXml)), rowsXml).toBe(expected);
    }
  });

  it("passes over a formula and the extensions of a row or a cell with all they hold", async () => {
    const part = data(
      '<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>7</v></c>' +
        '<c r="B1"><v>8</v><f>1+1</f><extLst><ext uri="x"><v>9</v></ext></extLst></c>' +
        '<extLst><ext uri="x"><c r="C1"><v>3</v></c></ext></extLst></row>',
    );
    expect(await rowsOf(part)).toEqual([[1, 7, 8]]);
  });

  it("reads around what else the worksheet holds, and sizes nothing from it", async () => {
    const part = worksheet(
      '<sheetPr><outlinePr summaryBelow="0"/></sheetPr><dimension ref="A1:XFD1048576"/>' +
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="16384" width="9"/></cols>' +
        '<sheetData><row r="1" spans="1:16384" hidden="1"><c r="A1"><v>1</v></c></row></sheetData>' +
        `<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>${alternate('<x14:foo xmlns:x14="urn:x14"/>')}<extLst><ext uri="x"/></extLst>`,
    );
    expect(await rowsOf(part, { limits: { cellPositions: 1 } })).toEqual([[1, 1]]);
    expect(await rowsOf(data(""))).toEqual([]);
  });

  it("refuses with INVALID_WORKBOOK a worksheet without sheet data, with two, or with sheet data or a row anywhere else", async () => {
    const row = '<row r="1"><c r="A1"><v>1</v></c></row>';
    const parts = [
      worksheet(""),
      worksheet("<sheetPr/>"),
      `<chartsheet xmlns="${MAIN}"><sheetData>${row}</sheetData></chartsheet>`,
      `<sheetData xmlns="${MAIN}">${row}</sheetData>`,
      worksheet(`<sheetData>${row}</sheetData><sheetData/>`),
      worksheet(alternate(`<sheetData>${row}</sheetData>`)),
      worksheet(`<sheetData/>${alternate(`<sheetData>${row}</sheetData>`)}`),
      worksheet(`${row}<sheetData/>`),
      worksheet(`<sheetData/>${row}`),
      worksheet(`<sheetPr>${row}</sheetPr><sheetData/>`),
      worksheet(`<sheetPr><sheetData>${row}</sheetData></sheetPr>`),
      worksheet(`<sheetPr><outlinePr/><sheetData>${row}</sheetData></sheetPr>`),
      data(`<foo/>${row}`),
      data(`<foo>${row}</foo>`),
      data(alternate(row)),
      data('<row r="1"><c r="A1"><f><sheetData/></f></c></row>'),
      data('<row r="1"><c r="A1"><f><x/><sheetData/></f></c></row>'),
      data('<row r="1"><extLst><ext uri="x"/><sheetData/></extLst></row>'),
      data(`<row r="1"><extLst><ext uri="x">${row.replace('r="1"', 'r="2"')}</ext></extLst></row>`),
    ];
    for (const part of parts) expect(await outcomeOf(part), part).toBe("INVALID_WORKBOOK");
  });

  it("reads SpreadsheetML in either family and under any prefix alike", async () => {
    const cells = '<c r="A1" t="inlineStr"><is><r><t>a</t></r></is></c><c r="B1"><v>2</v></c>';
    const prefixedCells = cells.replaceAll(/<(\/?)(\w)/g, "<$1x:$2");
    const row = `<row r="1">${cells}</row>`;
    const prefixed = `<x:row r="1">${prefixedCells}</x:row>`;
    for (const part of [
      data(row),
      `<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData>${row}</sheetData></worksheet>`,
      `<x:worksheet xmlns:x="${MAIN}"><x:sheetData>${prefixed}</x:sheetData></x:worksheet>`,
      `<x:worksheet xmlns:x="${MAIN}" xmlns="urn:other"><dimension/><x:sheetData>${prefixed}</x:sheetData></x:worksheet>`,
      data(`<row r="1" xmlns:x="${MAIN}">${prefixedCells}</row>`),
    ]) {
      expect(await outcomeOf(part), part).toEqual([[1, "a", 2]]);
    }
  });

  it("refuses an element of another namespace on the path to a value where one of an unknown name is refused, and gives none of its content", async () => {
    // The shapes of the stored cases, with what this reader gives, where LibreOffice and openpyxl
    // read them as the table of issue 105 says.
    const marker = '<row r="1"><c r="A1"><v>7</v></c></row>';
    const q = 'xmlns:q="urn:foreign"';
    const shapes: [shape: string, part: string, outcome: string][] = [
      [
        "sheetData in a foreign default namespace",
        worksheet(`<sheetData xmlns="urn:foreign">${marker}</sheetData>`),
        "INVALID_WORKBOOK",
      ],
      [
        "sheetData prefixed foreign",
        worksheet(`<q:sheetData ${q}>${marker}</q:sheetData>`),
        "INVALID_WORKBOOK",
      ],
      [
        "row prefixed foreign",
        data(`<q:row ${q} r="1"><c r="A1"><v>1</v></c></q:row>`),
        "INVALID_WORKBOOK",
      ],
      [
        "row in a foreign default namespace",
        data('<row xmlns="urn:foreign" r="1"><c r="A1"><v>1</v></c></row>'),
        "INVALID_WORKBOOK",
      ],
      [
        "c prefixed foreign",
        data(`${marker}<row r="2"><q:c ${q} r="A2"><v>1</v></q:c></row>`),
        "INVALID_CELL A2",
      ],
      [
        "c in a foreign default namespace",
        data(`${marker}<row r="2"><c xmlns="urn:foreign" r="A2"><v>1</v></c></row>`),
        "INVALID_CELL A2",
      ],
      [
        "v prefixed foreign",
        data(`${marker}<row r="2"><c r="A2"><q:v ${q}>1</q:v></c></row>`),
        "INVALID_CELL A2",
      ],
      [
        "v in a foreign default namespace",
        data(`${marker}<row r="2"><c r="A2"><v xmlns="urn:foreign">1</v></c></row>`),
        "INVALID_CELL A2",
      ],
      [
        "is prefixed foreign",
        data(`${marker}<row r="2"><c r="A2" t="inlineStr"><q:is ${q}><t>a</t></q:is></c></row>`),
        "INVALID_CELL A2",
      ],
      [
        "t of is prefixed foreign",
        data(`${marker}<row r="2"><c r="A2" t="inlineStr"><is><q:t ${q}>a</q:t></is></c></row>`),
        "INVALID_CELL A2",
      ],
      [
        "worksheet root in a foreign namespace",
        `<worksheet xmlns="urn:foreign"><sheetData>${marker}</sheetData></worksheet>`,
        "INVALID_WORKBOOK",
      ],
      [
        "worksheet root in no namespace",
        `<worksheet><sheetData>${marker}</sheetData></worksheet>`,
        "INVALID_WORKBOOK",
      ],
      [
        "c in main and a foreign c for the same place",
        data(`<row r="1"><c r="A1"><v>1</v></c><q:c ${q} r="B1"><v>2</v></q:c></row>`),
        "INVALID_CELL B1",
      ],
      [
        "an extLst of another namespace in a row",
        data(`<row r="1"><c r="A1"><v>1</v></c><q:extLst ${q}/></row>`),
        "INVALID_CELL B1",
      ],
      [
        "an f of another namespace in a cell",
        data(`<row r="1"><c r="A1"><q:f ${q}>1+1</q:f><v>2</v></c></row>`),
        "INVALID_CELL A1",
      ],
    ];
    for (const [shape, part, outcome] of shapes) expect(await outcomeOf(part), shape).toBe(outcome);
  });

  it("passes over an element of another namespace with what it holds where one of an unknown name is passed over", async () => {
    const marker = '<row r="1"><c r="A1"><v>7</v></c></row>';
    for (const part of [
      worksheet(
        `<q:sheetPr xmlns:q="urn:foreign"><q:row/></q:sheetPr><sheetData>${marker}</sheetData>`,
      ),
      worksheet(
        `<sheetData>${marker}</sheetData><foo xmlns="urn:foreign"><sheetData/><row/></foo>`,
      ),
      data(`<row r="1"><c r="A1"><v>7</v></c><extLst><q:c xmlns:q="urn:foreign"/></extLst></row>`),
    ]) {
      expect(await outcomeOf(part), part).toEqual([[1, 7]]);
    }
  });

  it("reads no attribute under a prefix as the one it names: a cell's reference, type and style, a row's number", async () => {
    const x = `xmlns:x="${MAIN}"`;
    expect(
      await outcomeOf(data(`<row r="1" ${x} x:r="5"><c r="A1" x:r="C1"><v>1</v></c></row>`)),
    ).toEqual([[1, 1]]);
    expect(
      await outcomeOf(data(`<row ${x} x:r="5"><c x:r="C1" x:t="s"><v>0</v></c></row>`)),
    ).toEqual([[1, 0]]);
    expect(
      await outcomeOf(data(`<row r="1" ${x}><c r="A1" x:s="1"><v>45000</v></c></row>`), {
        dates: [0, 1],
      }),
    ).toEqual([[1, 45000]]);
  });

  it("gives each row as it closes, before the rest of the part is read", async () => {
    let reads = 0;
    const chunks = [data('<row r="1"><c r="A1"><v>1</v></c></row><row r="2">'), "never read"];
    const { positions, rowText } = resolveLimits(undefined).rows();
    const visitor = worksheetVisitor({
      strings: [],
      dates: undefined,
      date1904: false,
      positions,
      rowText,
    });
    const bytes: ByteChunks = {
      read: () => Promise.resolve(encoder.encode(chunks[reads++] ?? "")),
      close() {},
    };
    const first = await scanXml(bytes, visitor).next();
    expect(first.value).toEqual({ number: 1, cells: [1] });
    expect(reads).toBe(1);
  });
});

const DAY = 86_400_000;
const utc = (iso: string) => new Date(`${iso}Z`);

/** What a number in a cell of style 1, a date style, reads as. */
async function datedOf(serial: string, date1904 = false): Promise<ReadValue | string> {
  const rows = await outcomeOf(cell(' s="1"', `<v>${serial}</v>`), { dates: [0, 1], date1904 });
  return typeof rows === "string" ? rows : (rows[0]?.[1] ?? null);
}

describe("dates", () => {
  it("gives a number whose style shows a date as the Date the writer reads it from, and keeps a serial the writer cannot write as a number", async () => {
    const serials1900: readonly (readonly [string, ReadValue])[] = [
      ["-1", -1],
      ["-0", -0],
      ["0", 0],
      ["0.5", 0.5],
      ["0.9999999999999", 0.9999999999999],
      ["1", utc("1900-01-01T00:00:00")],
      ["59.5", utc("1900-02-28T12:00:00")],
      ["59.9999999", utc("1900-02-28T23:59:59.991")],
      ["59.99999999999", 59.99999999999],
      ["60", 60],
      ["60.5", 60.5],
      ["60.9999999999999", 60.9999999999999],
      ["61", utc("1900-03-01T00:00:00")],
      ["45000.9999999999", utc("2023-03-16T00:00:00")],
      ["46343.52425925926", utc("2026-11-17T12:34:56")],
      ["2958465.9999999884", utc("9999-12-31T23:59:59.999")],
      ["2958465.9999999995", 2958465.9999999995],
      ["2958466", 2958466],
      ["1e308", 1e308],
    ];
    for (const [serial, value] of serials1900) {
      expect(await datedOf(serial), serial).toEqual(value);
    }
    const serials1904: readonly (readonly [string, ReadValue])[] = [
      ["-1", -1],
      ["-0.0000000001", -0.0000000001],
      ["-0", utc("1904-01-01T00:00:00")],
      ["0", utc("1904-01-01T00:00:00")],
      ["59.5", utc("1904-02-29T12:00:00")],
      ["60", utc("1904-03-01T00:00:00")],
      ["44881.52425925926", utc("2026-11-17T12:34:56")],
      ["2957003.9999999884", utc("9999-12-31T23:59:59.999")],
      ["2957003.9999999995", 2957003.9999999995],
      ["2957004", 2957004],
    ];
    for (const [serial, value] of serials1904) {
      expect(await datedOf(serial, true), `1904 ${serial}`).toEqual(value);
    }
    expect(Object.is(await datedOf("-0"), -0)).toBe(true);
  });

  it("gives only Dates the writer writes back to within half a millisecond of the serial, and keeps as numbers only the serials at or past an edge", async () => {
    const nearEdge = (edge: number) =>
      fc.double({ min: -2, max: 2, noNaN: true }).map((millis) => edge + millis / DAY);
    const system = fc.oneof(
      fc
        .oneof(
          ...[1, 60, 61, 45001, 2958466].map(nearEdge),
          fc.double({ min: -2, max: 3e6, noNaN: true }),
        )
        .map((serial) => ({ serial, date1904: false })),
      fc
        .oneof(...[0, 2957004].map(nearEdge), fc.double({ min: -2, max: 3e6, noNaN: true }))
        .map((serial) => ({ serial, date1904: true })),
    );
    await fc.assert(
      fc.asyncProperty(system, async ({ serial, date1904 }) => {
        const value = await datedOf(String(serial), date1904);
        if (typeof value === "number") {
          expect(value).toBe(Number(String(serial)));
          const [lowest, last] = date1904 ? [0, 2957004] : [1, 2958466];
          const justBelow = (edge: number) => serial < edge && (edge - serial) * DAY < 1;
          const leapDay = !date1904 && (justBelow(60) || (serial >= 60 && serial < 61));
          expect(serial < lowest || leapDay || justBelow(last) || serial >= last).toBe(true);
          return;
        }
        expect(value).toBeInstanceOf(Date);
        const written = serialFromEpochMillis((value as Date).getTime());
        expect(written).not.toBeNaN();
        expect(Math.abs(written - (date1904 ? 1462 : 0) - serial) * DAY).toBeLessThan(0.6);
      }),
      { numRuns: 4_000 },
    );
  });

  it("reads a number as a Date only by its style: text, a boolean and an error keep their type, and a formula's stored number is a date", async () => {
    const cells = [
      '<c r="A1" s="1" t="s"><v>0</v></c>',
      '<c r="B1" s="1" t="str"><v>45000</v></c>',
      '<c r="C1" s="1" t="b"><v>1</v></c>',
      '<c r="D1" s="1" t="e"><v>#N/A</v></c>',
      '<c r="E1" s="1" t="inlineStr"><is><t>45000</t></is></c>',
      '<c r="F1" s="1"><f>A1+1</f><v>45000.5</v></c>',
      '<c r="G1" s="1" t="n"><v>45000</v></c>',
      '<c r="H1" s="0"><v>45000</v></c>',
      '<c r="I1"><v>45000</v></c>',
    ];
    const part = data(`<row r="1">${cells.join("")}</row>`);
    expect(await rowsOf(part, { strings: ["45000"], dates: [0, 1] })).toEqual([
      [
        1,
        "45000",
        "45000",
        true,
        "#N/A",
        "45000",
        utc("2023-03-15T12:00:00"),
        utc("2023-03-15T00:00:00"),
        45000,
        45000,
      ],
    ]);
    expect(await rowsOf(part, { strings: ["45000"], dates: [1, 0] })).toEqual([
      [
        1,
        "45000",
        "45000",
        true,
        "#N/A",
        "45000",
        45000.5,
        45000,
        utc("2023-03-15T00:00:00"),
        utc("2023-03-15T00:00:00"),
      ],
    ]);
  });

  it("reads a date stored as text in the forms of XML Schema's date and dateTime, with white space around it and decimals rounded to the millisecond, as a Date whatever its style", async () => {
    const forms: readonly (readonly [string, Date])[] = [
      ["2026-11-17", utc("2026-11-17T00:00:00")],
      ["2026-11-17Z", utc("2026-11-17T00:00:00")],
      ["2026-11-17T12:34:56", utc("2026-11-17T12:34:56")],
      ["2026-11-17T12:34:56Z", utc("2026-11-17T12:34:56")],
      ["2026-11-17T12:34:56.5", utc("2026-11-17T12:34:56.500")],
      ["2026-11-17T12:34:56.12", utc("2026-11-17T12:34:56.120")],
      ["2026-11-17T12:34:56.123Z", utc("2026-11-17T12:34:56.123")],
      ["2024-02-29", utc("2024-02-29T00:00:00")],
      ["1900-01-01", utc("1900-01-01T00:00:00")],
      ["1900-02-28T23:59:59", utc("1900-02-28T23:59:59")],
      ["1900-03-01", utc("1900-03-01T00:00:00")],
      ["9999-12-31T23:59:59.999", utc("9999-12-31T23:59:59.999")],
      [" 2026-11-17", utc("2026-11-17T00:00:00")],
      ["2026-11-17 ", utc("2026-11-17T00:00:00")],
      ["\t\n 2026-11-17T12:34:56Z \n\t", utc("2026-11-17T12:34:56")],
      ["2026-11-17T12:34:56.1234", utc("2026-11-17T12:34:56.123")],
      ["2026-11-17T12:34:56.1235", utc("2026-11-17T12:34:56.124")],
      ["2026-11-17T12:34:56.0005", utc("2026-11-17T12:34:56.001")],
      ["2026-11-17T12:34:56.1234567", utc("2026-11-17T12:34:56.123")],
      ["2026-11-17T12:34:56.12349999999999999999Z", utc("2026-11-17T12:34:56.123")],
      ["2026-11-17T12:34:56.9999999Z", utc("2026-11-17T12:34:57.000")],
      [`2026-11-17T12:34:56.${"9".repeat(20)}`, utc("2026-11-17T12:34:57.000")],
      [`2026-11-17T12:34:56.${"1".repeat(400)}`, utc("2026-11-17T12:34:56.111")],
      ["2026-01-31T23:59:59.9995", utc("2026-02-01T00:00:00")],
      ["2026-12-31T23:59:59.99951", utc("2027-01-01T00:00:00")],
      ["1900-02-28T23:59:59.9995", utc("1900-03-01T00:00:00")],
      ["9999-12-31T23:59:59.9994999", utc("9999-12-31T23:59:59.999")],
    ];
    for (const style of ["", ' s="0"', ' s="1"', ' s="2"']) {
      for (const [text, date] of forms) {
        const part = cell(` t="d"${style}`, `<v>${text}</v>`);
        expect(await rowsOf(part, { dates: [0, 1, 0] }), `${style} ${text}`).toEqual([[1, date]]);
      }
    }
    expect(await rowsOf(cell(' t="d"', "<v>1904-01-01</v>"), { date1904: true })).toEqual([
      [1, utc("1904-01-01T00:00:00")],
    ]);
  });

  it("refuses a date stored as text in any other form, a day that does not exist and one the writer cannot write once rounded", async () => {
    const refused = [
      "2026-11-17T12:34:56+02:00",
      "2026-11-17T12:34:56-05:00",
      "2026-11-17T12:34:56+00:00",
      "2026-11-17+02:00",
      "2026-11-17T12:34:56.1234+01:00",
      "2026-11-17T12:34:56.",
      "2026-11-17T12:34:56,5",
      "2026-11-17T12:34",
      "2026-11-17T12",
      "2026-11-17T",
      "2026-11-17 12:34:56",
      "2026-11-17t12:34:56",
      "2026-11-17T12:34:56z",
      "12:34:56",
      "T12:34:56",
      "2026-11-17T24:00:00",
      "2026-11-17T12:60:00",
      "2026-11-17T12:34:60",
      "2026-11-17T1:2:3",
      "2026-11-17T12.34.56",
      "2026-11-17T12.34:56",
      "2026/11-17",
      "2026-11-1/",
      "2026-11-17T12:3/:56",
      "2026-11-17T12:34.56",
      "2026/11/17",
      "2026-11/17",
      "2026-11-17T1a:34:56",
      "2026-11-17T12:3a:56",
      "2026-11-17T12:34:5a",
      "2026-02-29",
      "2026-02-30",
      "2026-04-31",
      "1900-02-29",
      "2026-13-01",
      "2026-00-10",
      "2026-11-00",
      "1899-12-31",
      "1899-12-31T23:59:59.9995",
      "9999-12-31T23:59:59.9995",
      "9999-12-31T23:59:59.99999999999",
      "0001-01-01",
      "10000-01-01",
      "+2026-11-17",
      "-2026-11-17",
      "2026-1-7",
      "20261117",
      "2026-W47-2",
      "2026-321",
      " ",
      "2026-11-17 Z",
      "\u00a02026-11-17",
      "2026-11-17\u3000",
      "45000",
      "2026-11-17T12:34:56ZZ",
      "２０２６-11-17",
    ];
    for (const text of refused) {
      await expect(rowsOf(cell(' t="d"', `<v>${text}</v>`)), text).rejects.toMatchObject({
        code: "INVALID_CELL",
        cell: "A1",
        message: "a date stored as text is malformed or out of range",
      });
    }
  });

  it("reads a cell's style only in a number and only when given a table of styles, as where a style shows a date", async () => {
    const elsewhere: readonly (readonly [string, string, ReadValue])[] = [
      ["", "", null],
      [' t="n"', "", null],
      [' t="s"', "<v>0</v>", "text"],
      [' t="str"', "<v>45000</v>", "45000"],
      [' t="inlineStr"', "<is><t>45000</t></is>", "45000"],
      [' t="b"', "<v>1</v>", true],
      [' t="e"', "<v>#N/A</v>", "#N/A"],
      [' t="d"', "<v>2023-03-15</v>", utc("2023-03-15T00:00:00")],
    ];
    const valueOf = (rows: Rows) => rows[0]?.[1] ?? null;
    for (const style of [
      "2",
      "99",
      "4294967296",
      "1.0",
      "-1",
      " 1",
      "+1",
      "",
      "1e0",
      "0x1",
      "&#49;",
    ]) {
      for (const type of ["", ' t="n"']) {
        const number = cell(` s="${style}"${type}`, "<v>45000</v>");
        expect(await outcomeOf(number, { dates: [0, 1] }), style).toBe("INVALID_CELL A1");
        expect(await rowsOf(number), style).toEqual([[1, 45000]]);
      }
      for (const [type, inner, value] of elsewhere) {
        const part = cell(` s="${style}"${type}`, inner);
        const rows = await rowsOf(part, { strings: ["text"], dates: [0, 1] });
        expect(valueOf(rows), `${style} ${type}`).toEqual(value);
      }
    }
    for (const style of ["", ' s="0"', ' s="00"']) {
      expect(await rowsOf(cell(style, "<v>1</v>"), { dates: [0, 1] }), style).toEqual([[1, 1]]);
    }
  });
});

describe("the limits of one read of a worksheet", () => {
  it("charges cellPositions the positions a value adds to its row, the empty ones before it included, and nothing for a cell without one", async () => {
    const corner = data('<row r="1"><c r="XFD1"><v>1</v></c></row>');
    expect(await outcomeOf(corner, { limits: { cellPositions: 16_384 } })).toHaveLength(1);
    expect(await outcomeOf(corner, { limits: { cellPositions: 16_383 } })).toBe(
      "LIMIT_EXCEEDED cellPositions",
    );
    const rows = data(
      '<row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>2</v></c><c r="Z1"/></row><row r="9"><c r="B9"><v>3</v></c></row>',
    );
    expect(await outcomeOf(rows, { limits: { cellPositions: 5 } })).toHaveLength(2);
    expect(await outcomeOf(rows, { limits: { cellPositions: 4 } })).toBe(
      "LIMIT_EXCEEDED cellPositions",
    );
    const valueless = data(
      '<row r="1"><c r="XFD1" s="3"/></row><row r="2"><c r="XFD2"><f>1</f></c></row>',
    );
    expect(
      await outcomeOf(valueless, { limits: { cellPositions: 1 }, dates: [0, 0, 0, 1] }),
    ).toEqual([]);
  });

  it("charges rowCharacters the length of every text placed in a row, a shared string each time, and starts again for every row", async () => {
    const strings = ["abc"];
    const shared = (count: number) =>
      Array.from(
        { length: count },
        (_, at) => `<c r="${String.fromCharCode(65 + at)}1" t="s"><v>0</v></c>`,
      ).join("");
    const limits = { rowCharacters: 6 };
    expect(
      await outcomeOf(data(`<row r="1">${shared(2)}</row>`), { strings, limits }),
    ).toHaveLength(1);
    expect(await outcomeOf(data(`<row r="1">${shared(3)}</row>`), { strings, limits })).toBe(
      "LIMIT_EXCEEDED rowCharacters",
    );
    const rows = [1, 2, 3].map(
      (number) => `<row r="${number}">${shared(2).replaceAll('1"', `${number}"`)}</row>`,
    );
    expect(await outcomeOf(data(rows.join("")), { strings, limits })).toHaveLength(3);
    const kinds = [
      '<c r="A1" t="str"><v>abcd</v></c><c r="B1" t="e"><v>#N/A</v></c>',
      '<c r="A1" t="e"><v>#DIV/0!</v></c>',
    ];
    for (const cells of kinds) {
      expect(await outcomeOf(data(`<row r="1">${cells}</row>`), { limits }), cells).toBe(
        "LIMIT_EXCEEDED rowCharacters",
      );
    }
    const untexted =
      '<c r="A1"><v>123456789</v></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="str"><v>abcdef</v></c>';
    expect(await outcomeOf(data(`<row r="1">${untexted}</row>`), { limits })).toHaveLength(1);
  });

  it("charges an inline string to rowCharacters run by run as it is read, and once", async () => {
    const runs = (count: number) =>
      `<row r="1"><c r="A1" t="inlineStr"><is>${"<r><t>abc</t></r>".repeat(count)}</is></c></row>`;
    expect(await outcomeOf(data(runs(2)), { limits: { rowCharacters: 6 } })).toEqual([
      [1, "abcabc"],
    ]);
    expect(await outcomeOf(data(runs(3)), { limits: { rowCharacters: 6 } })).toBe(
      "LIMIT_EXCEEDED rowCharacters",
    );
    const phonetic =
      '<row r="1"><c r="A1" t="inlineStr"><is><t>abc</t><rPh sb="0" eb="1"><t>PHONETIC</t></rPh></is></c></row>';
    expect(await outcomeOf(data(phonetic), { limits: { rowCharacters: 3 } })).toEqual([[1, "abc"]]);
  });
});
