import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vite-plus/test";
import { type WriteOptions, type Workbook, writeXlsx } from "../src/index.js";
import { cellTypesOptions, cellTypesWorkbook, minimalWorkbook } from "./fixtures/workbooks.js";
import { readPart, unzip } from "./helpers/unzip.js";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

async function golden(
  name: string,
  workbook: Workbook,
  options?: WriteOptions,
): Promise<Uint8Array> {
  const bytes = await writeXlsx(workbook, options);
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
