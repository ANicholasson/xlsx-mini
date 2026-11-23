import fc from "fast-check";
import { describe, expect, it } from "vite-plus/test";
import { type ReadLimits, XlsxReadError, resolveLimits } from "../../errors/read.js";
import { type ByteChunks, readXml } from "../../xml/read.js";
import { createStringTable } from "../index.js";
import { type SharedStrings, decodeEscapes, sharedStringsVisitor } from "../read.js";

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

const sst = (items: string, attributes = "") => `<sst xmlns="${MAIN}"${attributes}>${items}</sst>`;

function tableOf(part: string, limits?: ReadLimits): Promise<SharedStrings> {
  return readXml(bytesOf(part), sharedStringsVisitor(resolveLimits(limits).book.strings));
}

async function outcomeOf(part: string, limits?: ReadLimits): Promise<SharedStrings | string> {
  try {
    return await tableOf(part, limits);
  } catch (error) {
    if (error instanceof XlsxReadError) return `${error.code} ${error.limit ?? ""}`.trim();
    throw error;
  }
}

describe("the shared-string table", () => {
  it("reads each entry as LibreOffice 26.8 shows it: runs joined, phonetic runs and run properties left out, nothing trimmed", async () => {
    const entries = {
      "<si><t>plain</t></si>": "plain",
      "<si><t>a</t><r><t>b</t></r></si>": "ab",
      '<si><r><t>x</t></r><rPh sb="0" eb="1"><t>PHON</t></rPh><phoneticPr fontId="1"/></si>': "x",
      "<si><r><rPr><b/></rPr><t>run</t></r><r><t>two</t></r></si>": "runtwo",
      "<si><r><rPr><t>inrpr</t></rPr><t>r</t></r></si>": "r",
      "<si><t>a</t><r><rPr/></r><r><t>c</t></r></si>": "ac",
      "<si><r><t>_x00</t></r><r><t>41_</t></r></si>": "_x0041_",
      "<si><t>_x0041_</t></si>": "A",
      "<si><t>_x00e6_|_x00E6_|_X0041_|_x41_|_xZZZZ_</t></si>": "æ|æ|_X0041_|_x41_|_xZZZZ_",
      "<si><t>a_xD83D__xDE00_b</t></si>": "a\u{1F600}b",
      "<si><t>_x0041_x0042_</t></si>": "Ax0042_",
      "<si><t>_x_x0041_</t></si>": "_xA",
      "<si><t>_x005F_x0041_</t></si>": "_x0041_",
      "<si><t>__x0078_0041_</t></si>": "_x0041_",
      "<si/>": "",
      "<si><t></t></si>": "",
      "<si>abc</si>": "",
      '<si><t xml:space="preserve">  sp  </t></si>': "  sp  ",
      "<si><t>  nsp  </t></si>": "  nsp  ",
      "<si><t>a&#13;b\r\nc&amp;</t></si>": "a\rb\nc&",
    };
    expect(await tableOf(sst(Object.keys(entries).join("")))).toEqual(Object.values(entries));
  });

  it("reads back every text the writer's table holds, escapes, carriage returns and markup included", async () => {
    const alphabet = fc.constantFrom(
      "_",
      "x",
      "X",
      "0",
      "4",
      "1",
      "a",
      "F",
      "f",
      "9",
      " ",
      "\r",
      "\n",
      "\t",
      "&",
      "<",
      ">",
      '"',
      "'",
      "é",
      "€",
      "\u{1F389}",
    );
    await fc.assert(
      fc.asyncProperty(
        fc.array(fc.array(alphabet, { maxLength: 24 }), { maxLength: 12 }),
        async (texts) => {
          const table = createStringTable();
          const written = texts.map((characters) => characters.join(""));
          const indexes = written.map((text) => table.intern(text));
          const read = await tableOf(table.render());
          expect(indexes.map((index) => read[index])).toEqual(written);
        },
      ),
    );
  });

  it("refuses with INVALID_WORKBOOK a root that is not sst, and anything under it but an entry or extensions without one", async () => {
    const parts = [
      `<sheet>${"<si><t>a</t></si>"}</sheet>`,
      sst("<si><t>a</t></si><item><t>b</t></item>"),
      sst(
        `<si><t>zero</t></si><mc:AlternateContent xmlns:mc="${COMPATIBILITY}"><mc:Choice Requires="x"><si><t>hidden</t></si></mc:Choice></mc:AlternateContent><si><t>one</t></si>`,
      ),
      sst("<sst><si><t>a</t></si></sst>"),
      `<extLst xmlns="${MAIN}"><ext uri="x"/></extLst>`,
      sst(
        '<si><t>a</t></si><extLst><ext uri="x"><si><t>b</t></si></ext></extLst><si><t>c</t></si>',
      ),
    ];
    for (const part of parts) expect(await outcomeOf(part), part).toBe("INVALID_WORKBOOK");
    expect(
      await outcomeOf(
        sst('<si><t>a</t></si><extLst><ext uri="x"><x:t>b</x:t></ext></extLst><si><t>c</t></si>'),
      ),
    ).toEqual(["a", "c"]);
  });

  it("refuses an entry holding what LibreOffice passes over or reads in part: an unknown element, one in a run, a child of a text", async () => {
    const items = [
      "<si><foo><t>hidden</t></foo><t>shown</t></si>",
      "<si><r><t>a</t><foo/></r></si>",
      "<si><r><r><t>a</t></r></r></si>",
      "<si><t>a<b>bold</b>c</t></si>",
      "<si><r><t>a<b/></t></r></si>",
      "<si><r><t>a<t>b</t>c</t></r></si>",
      "<si><r><t>a<rPr/>b</t></r></si>",
      "<si><rPr/><t>a</t></si>",
      "<si><r><rPh><t>a</t></rPh></r></si>",
    ];
    for (const item of items) expect(await outcomeOf(sst(item)), item).toBe("INVALID_WORKBOOK");
  });

  it("refuses an entry that LibreOffice and openpyxl read in two ways: a second text of its own or of a run, its own text after a run", async () => {
    const items = [
      "<si><t>t1</t><t>t2</t></si>",
      "<si><r><t>a</t><t>b</t></r><r><t>c</t></r></si>",
      "<si><r><t>a</t></r><t>b</t></si>",
      "<si><r>abc</r><t>d</t></si>",
    ];
    for (const item of items) expect(await outcomeOf(sst(item)), item).toBe("INVALID_WORKBOOK");
    expect(await outcomeOf(sst("<si><t>a</t></si><si><t>b</t><r><t>c</t></r></si>"))).toEqual([
      "a",
      "bc",
    ]);
  });

  it("charges every entry to sharedStrings before it keeps it, and never reads the count the table states", async () => {
    const three = sst("<si><t>a</t></si><si/><si><t>c</t></si>", ' count="1" uniqueCount="1"');
    expect(await outcomeOf(three, { sharedStrings: 3 })).toEqual(["a", "", "c"]);
    expect(await outcomeOf(three, { sharedStrings: 2 })).toBe("LIMIT_EXCEEDED sharedStrings");
    expect(
      await outcomeOf(sst("<si><t>a</t></si>", ' count="9999999999" uniqueCount="9999999999"'), {
        sharedStrings: 1,
      }),
    ).toEqual(["a"]);
  });
});

describe("decodeEscapes", () => {
  it("decodes the forms Excel and LibreOffice decode, from left to right, and returns text without one as it is", () => {
    const text = "no escape here";
    expect(decodeEscapes(text)).toBe(text);
    expect(decodeEscapes("__x0078_0041_")).toBe("_x0041_");
    expect(decodeEscapes("a_x000D_b_x000a_")).toBe("a\rb\n");
    expect(decodeEscapes("_x0041_x0042_")).toBe("Ax0042_");
    expect(decodeEscapes("_x_x0041_")).toBe("_xA");
    expect(decodeEscapes("_x0041")).toBe("_x0041");
    const almost = "_x004G_ _x00@0_ _x00`0_ _x00/0_ _x00:0_ _x0041";
    expect(decodeEscapes(almost)).toBe(almost);
    expect(decodeEscapes("_xFFFF__xffff_")).toBe("￿￿");
  });
});
