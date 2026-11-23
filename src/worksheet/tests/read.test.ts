import { describe, expect, it } from "vite-plus/test";
import { type ReadLimits, XlsxReadError, resolveLimits } from "../../errors/read.js";
import { type ByteChunks, scanXml } from "../../xml/read.js";
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
/** One cell in A1 of one row. */
const cell = (attributes: string, inner: string) =>
  data(`<row r="1"><c r="A1"${attributes}>${inner}</c></row>`);

type Rows = readonly (readonly [number, ...ReadValue[]])[];

interface Reading {
  readonly strings?: readonly string[];
  readonly limits?: ReadLimits;
}

async function rowsOf(part: string, { strings = [], limits }: Reading = {}): Promise<Rows> {
  const { positions, rowText } = resolveLimits(limits).rows();
  const rows: (readonly [number, ...ReadValue[]])[] = [];
  const visitor = worksheetVisitor({ strings, positions, rowText });
  for await (const { number, cells } of scanXml(bytesOf(part), visitor)) {
    rows.push([number, ...cells]);
  }
  return rows;
}

/** The rows, or the code of the refusal with its cell or its limit. */
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
    expect(await rowsOf(sparse)).toEqual([
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

  it("reads a cell that stores nothing as empty, whatever its type, and an empty v as empty text where the type is text", async () => {
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
    for (const type of [' t="e"', ' t="str"']) {
      for (const inner of ["<v></v>", "<v/>", '<f>""</f><v xml:space="preserve"></v>']) {
        expect(await rowsOf(stored(type, inner)), `${type} ${inner}`).toEqual([[1, null, 2]]);
      }
    }
    for (const type of ["", ' t="n"', ' t="s"', ' t="b"', ' t="d"']) {
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

  it("refuses an unknown type, and a date stored as text until dates are read", async () => {
    const date = await outcomeOf(cell(' t="d"', "<v>2024-01-15T12:00:00</v>"));
    expect(date).toBe("INVALID_CELL A1");
    await expect(rowsOf(cell(' t="d"', "<v>2024-01-15</v>"))).rejects.toThrow(
      "a date stored as text is not read yet",
    );
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
    ];
    for (const reference of references) {
      const part = data(`<row r="1"><c r="${reference}"><v>1</v></c></row>`);
      const written = reference.replace("&#49;", "1");
      expect(await outcomeOf(part), reference).toBe(
        ["INVALID_CELL", written].filter(Boolean).join(" "),
      );
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
        `<mergeCells count="1"><mergeCell ref="A1:B2"/></mergeCells>${alternate("<x14:foo/>")}<extLst><ext uri="x"/></extLst>`,
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
      data(`<foo/>${row}`),
      data(`<foo>${row}</foo>`),
      data(alternate(row)),
      data('<row r="1"><c r="A1"><f><sheetData/></f></c></row>'),
      data(`<row r="1"><extLst><ext uri="x">${row.replace('r="1"', 'r="2"')}</ext></extLst></row>`),
    ];
    for (const part of parts) expect(await outcomeOf(part), part).toBe("INVALID_WORKBOOK");
  });

  it("gives each row as it closes, before the rest of the part is read", async () => {
    let reads = 0;
    const chunks = [data('<row r="1"><c r="A1"><v>1</v></c></row><row r="2">'), "never read"];
    const { positions, rowText } = resolveLimits(undefined).rows();
    const visitor = worksheetVisitor({ strings: [], positions, rowText });
    const bytes: ByteChunks = {
      read: () => Promise.resolve(encoder.encode(chunks[reads++] ?? "")),
      close() {},
    };
    const first = await scanXml(bytes, visitor).next();
    expect(first.value).toEqual({ number: 1, cells: [1] });
    expect(reads).toBe(1);
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
    expect(await outcomeOf(valueless, { limits: { cellPositions: 1 } })).toEqual([]);
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
