import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it, vi } from "vite-plus/test";
import { type WriteOptions, type Workbook, writeXlsx } from "../src/index.js";
import {
  cellTypesOptions,
  cellTypesWorkbook,
  minimalWorkbook,
  multiSheetWorkbook,
} from "./fixtures/workbooks.js";
import { readPart, unzip } from "./helpers/unzip.js";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

/** Golden fixtures are written with STORE entries: byte-deterministic on every platform. */
const STORE: WriteOptions = { compression: "none" };

async function golden(
  name: string,
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  const bytes = await writeXlsx(workbook, { ...options, ...STORE });
  const file = join(fixturesDir, name);
  if (process.env["UPDATE_FIXTURES"] === "1" || !existsSync(file)) {
    writeFileSync(file, bytes);
  }
  const expected = new Uint8Array(readFileSync(file));
  expect(bytes).toEqual(expected);
  return bytes;
}

describe("minimal workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("minimal.xlsx", minimalWorkbook);
  });

  it("contains exactly the parts a one-sheet workbook needs, content types first", async () => {
    const bytes = await writeXlsx(minimalWorkbook);
    expect(unzip(bytes).map((e) => e.name)).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/_rels/workbook.xml.rels",
      "xl/styles.xml",
      "xl/sharedStrings.xml",
      "xl/worksheets/sheet1.xml",
    ]);
  });

  it("declares every part in the content types and links the sheet from the workbook", async () => {
    const bytes = await writeXlsx(minimalWorkbook);
    const contentTypes = readPart(bytes, "[Content_Types].xml");
    for (const part of [
      "/xl/workbook.xml",
      "/xl/worksheets/sheet1.xml",
      "/xl/styles.xml",
      "/xl/sharedStrings.xml",
    ]) {
      expect(contentTypes).toContain(`PartName="${part}"`);
    }
    expect(readPart(bytes, "xl/workbook.xml")).toContain(
      '<sheet name="Sheet1" sheetId="1" r:id="rId1"/>',
    );
    expect(readPart(bytes, "xl/_rels/workbook.xml.rels")).toContain(
      'Target="worksheets/sheet1.xml"',
    );
    expect(readPart(bytes, "_rels/.rels")).toContain('Target="xl/workbook.xml"');
  });

  it("round-trips every string through the shared-strings table, escaped", async () => {
    const bytes = await writeXlsx(minimalWorkbook);
    const sst = readPart(bytes, "xl/sharedStrings.xml");
    expect(sst).toContain("<t>Ola &lt;Nordmann&gt; &amp; co</t>");
    expect(sst).toContain("<t>Ærlig «ekte» 🎉</t>");
    expect(sst).toContain('<t xml:space="preserve"> padded </t>');
    expect(sst).toContain("<t></t>");
    expect(sst).toContain('count="10" uniqueCount="10"');
  });

  it("writes numbers inline with full precision", async () => {
    const sheet = readPart(await writeXlsx(minimalWorkbook), "xl/worksheets/sheet1.xml");
    expect(sheet).toContain('<c r="C2"><v>1234.5</v></c>');
    expect(sheet).toContain('<c r="C3"><v>-42</v></c>');
    expect(sheet).toContain('<c r="C4"><v>1e+21</v></c>');
    expect(sheet).toContain('<c r="C5"><v>0</v></c>');
  });
});

describe("cell types workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("cell-types.xlsx", cellTypesWorkbook, cellTypesOptions);
  });

  it("writes every supported kind of value", async () => {
    const sheet = readPart(
      await writeXlsx(cellTypesWorkbook, cellTypesOptions),
      "xl/worksheets/sheet1.xml",
    );
    expect(sheet).toContain('<c r="B3"><v>3.14</v></c>');
    expect(sheet).toContain('<c r="B4" t="b"><v>1</v></c>');
    expect(sheet).toContain('<c r="B5" t="b"><v>0</v></c>');
    expect(sheet).toContain('<c r="B6" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B7" s="2"><v>46330</v></c>');
    expect(sheet).toContain('<c r="B8" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B9" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B10" s="1"><v>46330.57326388889</v></c>');
    expect(sheet).toContain('<c r="B11"><v>42</v></c>');
    expect(sheet).not.toContain('r="B12"');
    expect(sheet).toContain('<c r="B15"><v>12345678901234</v></c>');
  });

  it("strips control characters and keeps newlines in text", async () => {
    const sst = readPart(
      await writeXlsx(cellTypesWorkbook, cellTypesOptions),
      "xl/sharedStrings.xml",
    );
    expect(sst).toContain("<t>abc</t>");
    expect(sst).toContain("<t>line1\nline2</t>");
  });

  it("declares the date styles the cells reference", async () => {
    const styles = readPart(await writeXlsx(cellTypesWorkbook, cellTypesOptions), "xl/styles.xml");
    expect(styles).toContain('<cellXfs count="3">');
    expect(styles).toContain('numFmtId="14"');
    expect(styles).toContain('numFmtId="22"');
  });
});

describe("compression", () => {
  afterEach(() => vi.unstubAllGlobals());

  const partsOf = (bytes: Uint8Array) =>
    unzip(bytes).map((e) => [e.name, e.method, new TextDecoder().decode(e.data)] as const);

  it("is on by default and decompresses part by part to the STORE output", async () => {
    const stored = await writeXlsx(minimalWorkbook, STORE);
    const deflated = await writeXlsx(minimalWorkbook);
    expect(unzip(deflated).every((e) => e.method === 8)).toBe(true);
    expect(partsOf(deflated).map(([name, , text]) => [name, text])).toEqual(
      partsOf(stored).map(([name, , text]) => [name, text]),
    );
    expect(deflated.byteLength).toBeLessThan(stored.byteLength * 0.6);
  });

  it("matches the committed deflate fixture part by part (bytes differ by compressor)", async () => {
    const file = join(fixturesDir, "minimal-deflate.xlsx");
    if (process.env["UPDATE_FIXTURES"] === "1" || !existsSync(file)) {
      writeFileSync(file, await writeXlsx(minimalWorkbook));
    }
    const committed = new Uint8Array(readFileSync(file));
    expect(unzip(committed).every((e) => e.method === 8)).toBe(true);
    expect(partsOf(committed).map(([name, , text]) => [name, text])).toEqual(
      partsOf(await writeXlsx(minimalWorkbook, STORE)).map(([name, , text]) => [name, text]),
    );
  });

  it("falls back to STORE, byte-identical to the golden file, when the runtime cannot deflate", async () => {
    vi.stubGlobal("CompressionStream", undefined);
    const bytes = await writeXlsx(minimalWorkbook);
    expect(bytes).toEqual(new Uint8Array(readFileSync(join(fixturesDir, "minimal.xlsx"))));
  });
});

describe("multi-sheet workbook (golden)", () => {
  it("is byte-identical to the committed fixture", async () => {
    await golden("multi-sheet.xlsx", multiSheetWorkbook);
  });

  it("has one worksheet part, one relationship and one content-type override per sheet, in tab order", async () => {
    const bytes = await writeXlsx(multiSheetWorkbook, STORE);
    const names = unzip(bytes).map((e) => e.name);
    expect(names.filter((n) => n.startsWith("xl/worksheets/"))).toEqual([
      "xl/worksheets/sheet1.xml",
      "xl/worksheets/sheet2.xml",
      "xl/worksheets/sheet3.xml",
    ]);
    const contentTypes = readPart(bytes, "[Content_Types].xml");
    expect(contentTypes.match(/worksheets\/sheet\d\.xml/g)).toEqual([
      "worksheets/sheet1.xml",
      "worksheets/sheet2.xml",
      "worksheets/sheet3.xml",
    ]);
    const relationships = readPart(bytes, "xl/_rels/workbook.xml.rels");
    for (let i = 1; i <= 3; i++) {
      expect(relationships).toContain(
        `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`,
      );
    }
    expect(relationships).toContain(
      'Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"',
    );
    expect(relationships).toContain(
      'Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"',
    );
  });

  it("escapes sheet names in the workbook part and keeps their order", async () => {
    const workbook = readPart(await writeXlsx(multiSheetWorkbook, STORE), "xl/workbook.xml");
    expect(workbook).toContain(
      '<sheets><sheet name="Orders &amp; Returns" sheetId="1" r:id="rId1"/><sheet name="Ærlig &lt;oversikt&gt;" sheetId="2" r:id="rId2"/><sheet name="Empty" sheetId="3" r:id="rId3"/></sheets>',
    );
  });

  it("shares one strings table across sheets", async () => {
    const bytes = await writeXlsx(multiSheetWorkbook, STORE);
    const sst = readPart(bytes, "xl/sharedStrings.xml");
    expect(sst.match(/<t>shipped<\/t>/g)).toHaveLength(1);
    expect(sst).toContain('count="10" uniqueCount="7"');
    expect(readPart(bytes, "xl/worksheets/sheet3.xml")).toContain("<sheetData/>");
  });

  it("rejects a duplicate sheet name through the public function, at the second sheet", async () => {
    const duplicate = {
      sheets: [multiSheetWorkbook.sheets[0]!, { name: "orders & RETURNS", rows: [] }],
    };
    await expect(writeXlsx(duplicate)).rejects.toMatchObject({
      code: "DUPLICATE_SHEET_NAME",
      path: "sheets[1].name",
    });
  });
});
